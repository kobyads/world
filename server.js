import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir, readFile, readdir, rename, rm, stat, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, extname, join, sep } from 'node:path';
import { lookup } from 'node:dns/promises';
import { isIP } from 'node:net';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { load } from 'cheerio';
import { Document, HeadingLevel, ImageRun, Packer, Paragraph, TextRun } from 'docx';
import robotsParser from 'robots-parser';
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import sharp from 'sharp';

const directory = dirname(fileURLToPath(import.meta.url));
const jobRootDirectory = process.env.LOCALAPPDATA
  ? join(process.env.LOCALAPPDATA, 'Temp', 'opencode', 'site-to-word')
  : join(directory, '.data');
const jobDataDirectory = join(jobRootDirectory, 'jobs');
const port = Number(process.env.PORT || 3000);
const MAX_HTML_BYTES = 5 * 1024 * 1024;
const MAX_REQUEST_BYTES = 10 * 1024;
const MAX_REDIRECTS = 5;
const MAX_TOTAL_TEXT = 100 * 1024 * 1024;
const MAX_SOURCE_IMAGE_BYTES = 12 * 1024 * 1024;
const MAX_ARTICLE_IMAGE_BYTES = 32 * 1024 * 1024;
const MAX_ARCHIVE_DOCUMENT_BYTES = 0xf0000000;
const MAX_CRAWL_PAGES = 10_000;
const PAGE_CONCURRENCY = 4;
const IMAGE_CONCURRENCY = 3;
const KULTPROFITA_CONCURRENCY = 6;
const DOCX_CONCURRENCY = 4;
const MAX_DISCOVERY_DOCUMENTS = 8;
const CRAWLER_NAME = 'SiteToWord';
const jobs = new Map();

function isPrivateIPv4(address) {
  const parts = address.split('.').map(Number);
  if (parts.length !== 4 || parts.some((part) => part < 0 || part > 255)) return true;
  const [a, b] = parts;

  return a === 0 || a === 10 || a === 127 || a >= 224
    || (a === 100 && b >= 64 && b <= 127)
    || (a === 169 && b === 254)
    || (a === 172 && b >= 16 && b <= 31)
    || (a === 192 && (b === 0 || b === 168))
    || (a === 198 && (b === 18 || b === 19));
}

function isPrivateIPv6(address) {
  const normalized = address.toLowerCase().split('%')[0];
  if (normalized.startsWith('::ffff:')) {
    const mapped = normalized.slice(7);
    if (isIP(mapped) === 4) return isPrivateIPv4(mapped);
    return true;
  }

  return normalized === '::'
    || normalized === '::1'
    || normalized.startsWith('fc')
    || normalized.startsWith('fd')
    || /^fe[89ab]/.test(normalized)
    || normalized.startsWith('ff');
}

function isPrivateAddress(address) {
  const version = isIP(address);
  if (version === 4) return isPrivateIPv4(address);
  if (version === 6) return isPrivateIPv6(address);
  return true;
}

async function validatePublicUrl(value) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error('Введите корректную ссылку, например https://example.com');
  }

  if (!['http:', 'https:'].includes(url.protocol)) {
    throw new Error('Поддерживаются только ссылки http:// и https://');
  }
  if (url.username || url.password) {
    throw new Error('Ссылки с логином и паролем не поддерживаются');
  }
  if (url.port && !['80', '443'].includes(url.port)) {
    throw new Error('Разрешены только стандартные порты 80 и 443');
  }

  const hostname = url.hostname.toLowerCase().replace(/^\[|\]$/g, '');
  if (hostname === 'localhost' || hostname.endsWith('.localhost') || hostname.endsWith('.local')) {
    throw new Error('Ссылки на локальные адреса не поддерживаются');
  }

  if (isIP(hostname)) {
    if (isPrivateAddress(hostname)) throw new Error('Ссылки на локальные адреса не поддерживаются');
    return url;
  }

  let addresses;
  try {
    addresses = await lookup(hostname, { all: true, verbatim: true });
  } catch {
    throw new Error('Не удалось найти сайт. Проверьте ссылку и подключение к интернету');
  }
  if (!addresses.length || addresses.some(({ address }) => isPrivateAddress(address))) {
    throw new Error('Ссылки на локальные адреса не поддерживаются');
  }

  return url;
}

function checkRobots(url, robots) {
  if (robots?.isAllowed(url.toString(), CRAWLER_NAME) === false) {
    throw new Error('Правила robots.txt запрещают обработку указанной страницы');
  }
}

async function fetchPage(initialUrl, allowedOrigin = initialUrl.origin, robots = null, acceptXml = false) {
  let url = initialUrl;
  for (let redirectCount = 0; redirectCount <= MAX_REDIRECTS; redirectCount += 1) {
    checkRobots(url, robots);
    const response = await fetch(url, {
      redirect: 'manual',
      signal: AbortSignal.timeout(15_000),
      headers: {
        'User-Agent': 'SiteToWord/1.0 (+text extraction)',
        Accept: acceptXml
          ? 'text/html,application/xhtml+xml,application/rss+xml,application/atom+xml,application/xml,text/xml'
          : 'text/html,application/xhtml+xml',
      },
    });

    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get('location');
      await response.body?.cancel();
      if (!location || redirectCount === MAX_REDIRECTS) {
        throw new Error('Сайт перенаправляет слишком много раз');
      }
      url = await validatePublicUrl(new URL(location, url).toString());
      if (url.origin !== allowedOrigin) {
        throw new Error('Перенаправление на другой сайт остановлено');
      }
      continue;
    }

    if (!response.ok) {
      await response.body?.cancel();
      throw new Error(`Сайт ответил ошибкой: HTTP ${response.status}`);
    }
    const contentType = response.headers.get('content-type') || '';
    const isXml = acceptXml && /^(?:text\/xml|application\/(?:xml|rss\+xml|atom\+xml))(?:\s*;|$)/i.test(contentType);
    if (!isXml && !contentType.includes('text/html') && !contentType.includes('application/xhtml+xml')) {
      await response.body?.cancel();
      throw new Error('По этой ссылке находится не HTML-страница');
    }

    const declaredSize = Number(response.headers.get('content-length') || 0);
    if (declaredSize > MAX_HTML_BYTES) {
      await response.body?.cancel();
      throw new Error('Страница слишком большая для обработки');
    }

    const reader = response.body?.getReader();
    if (!reader) throw new Error('Не удалось прочитать страницу');
    const chunks = [];
    let totalBytes = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      totalBytes += value.byteLength;
      if (totalBytes > MAX_HTML_BYTES) {
        await reader.cancel();
        throw new Error('Страница слишком большая для обработки');
      }
      chunks.push(value);
    }

    const body = Buffer.concat(chunks);
    const declaredCharset = contentType.match(/charset\s*=\s*["']?([^;"'\s]+)/i)?.[1];
    const initialMarkup = body.subarray(0, 4096).toString('latin1');
    const metaCharset = initialMarkup.match(/<meta[^>]+charset\s*=\s*["']?([^"'\s/>]+)/i)?.[1]
      || initialMarkup.match(/<meta[^>]+content=["'][^"']*charset\s*=\s*([^\s;"']+)/i)?.[1];
    const charset = declaredCharset || metaCharset || 'utf-8';
    let html;
    try {
      html = new TextDecoder(charset).decode(body);
    } catch {
      html = body.toString('utf8');
    }
    return { html, url: url.toString(), isXml };
  }
  throw new Error('Не удалось загрузить страницу');
}

async function fetchJsonPageOnce(initialUrl, allowedOrigin, robots) {
  let url = initialUrl;
  for (let redirectCount = 0; redirectCount <= MAX_REDIRECTS; redirectCount += 1) {
    checkRobots(url, robots);
    const response = await fetch(url, {
      redirect: 'manual',
      signal: AbortSignal.timeout(15_000),
      headers: {
        'User-Agent': 'SiteToWord/1.0 (+text extraction)',
        Accept: 'application/json',
      },
    });
    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get('location');
      await response.body?.cancel();
      if (!location || redirectCount === MAX_REDIRECTS) throw new Error('Слишком много перенаправлений API');
      url = await validatePublicUrl(new URL(location, url).toString());
      if (url.origin !== allowedOrigin) throw new Error('Перенаправление API на другой сайт остановлено');
      continue;
    }
    if (!response.ok) {
      const error = new Error(`API сайта ответил ошибкой: HTTP ${response.status}`);
      error.status = response.status;
      error.retryAfter = response.headers.get('retry-after');
      await response.body?.cancel();
      throw error;
    }
    const contentType = response.headers.get('content-type') || '';
    if (!contentType.includes('application/json')) {
      await response.body?.cancel();
      throw new Error('API вернул ответ не в JSON');
    }
    const body = await readLimitedText(response, 10 * 1024 * 1024);
    try {
      return JSON.parse(body);
    } catch {
      throw new Error('API сайта вернул некорректный JSON');
    }
  }
  throw new Error('Не удалось загрузить данные API');
}

async function fetchJsonPage(initialUrl, allowedOrigin, robots = null) {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      return await fetchJsonPageOnce(initialUrl, allowedOrigin, robots);
    } catch (error) {
      const retryable = error.name === 'TimeoutError'
        || error.name === 'TypeError'
        || error.status === 429
        || error.status >= 500;
      if (!retryable || attempt === 2) throw error;
      const retryAfter = Number(error.retryAfter || 0) * 1000;
      await wait(Math.max(retryAfter, 350 * (attempt + 1)));
    }
  }
  throw new Error('Не удалось загрузить данные API');
}

async function readRobots(origin) {
  const robotsUrl = new URL('/robots.txt', origin);
  const safeUrl = await validatePublicUrl(robotsUrl.toString());
  const response = await fetch(safeUrl, {
    redirect: 'manual',
    signal: AbortSignal.timeout(10_000),
    headers: { 'User-Agent': 'SiteToWord/1.0 (+text extraction)' },
  });

  if (response.status === 404) return robotsParser(robotsUrl.toString(), '');
  if ([301, 302, 303, 307, 308].includes(response.status)) {
    const location = response.headers.get('location');
    if (!location || new URL(location, safeUrl).origin !== origin) {
      throw new Error('Перенаправление robots.txt на другой сайт остановлено');
    }
    const redirectedUrl = await validatePublicUrl(new URL(location, safeUrl).toString());
    const redirectedResponse = await fetch(redirectedUrl, {
      redirect: 'error',
      signal: AbortSignal.timeout(10_000),
      headers: { 'User-Agent': 'SiteToWord/1.0 (+text extraction)' },
    });
    if (!redirectedResponse.ok) throw new Error('Не удалось проверить правила robots.txt');
    return robotsParser(robotsUrl.toString(), await readLimitedText(redirectedResponse, 512 * 1024));
  }
  if (!response.ok) throw new Error('Не удалось проверить правила robots.txt');
  return robotsParser(robotsUrl.toString(), await readLimitedText(response, 512 * 1024));
}

async function readLimitedText(response, maxBytes) {
  const reader = response.body?.getReader();
  if (!reader) return '';
  const chunks = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel();
      throw new Error('Файл robots.txt слишком большой');
    }
    chunks.push(value);
  }
  return Buffer.concat(chunks).toString('utf8');
}

function normalizePublishedDate(candidate) {
  if (!candidate) return '';
  const value = String(candidate).trim();
  const datePart = value.match(/^(\d{4}-\d{2}-\d{2})/)?.[1]
    || value.match(/^(\d{2})\.(\d{2})\.(\d{4})$/)?.slice(1).reverse().join('-');
  const parsed = new Date(datePart ? `${datePart}T00:00:00.000Z` : value);
  return Number.isNaN(parsed.getTime()) ? '' : parsed.toISOString().slice(0, 10);
}

function cleanXmlText(value) {
  return String(value).replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu, '');
}

function extractPage(html, pageUrl = 'https://example.invalid') {
  const $ = load(html, { scriptingEnabled: false });
  let publishedDate = '';
  let structuredArticle;
  const structuredLinks = [];

  const dateCandidates = [
    $('meta[property="article:published_time"]').attr('content'),
    $('meta[name="datePublished"]').attr('content'),
    $('meta[itemprop="datePublished"]').attr('content'),
    $('time[datetime]').first().attr('datetime'),
    $('[itemprop="datePublished"]').first().text(),
  ];
  $('script[type="application/ld+json"]').each((_, script) => {
    try {
      const data = JSON.parse($(script).text());
      const collectMetadata = (value) => {
        if (Array.isArray(value)) {
          value.forEach(collectMetadata);
        } else if (value && typeof value === 'object') {
          if (typeof value.datePublished === 'string') dateCandidates.push(value.datePublished);
          const types = [value['@type']].flat();
          if (types.some((type) => typeof type === 'string' && /(?:Article|BlogPosting)$/i.test(type))) structuredArticle ||= value;
          if (Array.isArray(value.itemListElement)) {
            for (const entry of value.itemListElement) {
              structuredLinks.push(entry?.item?.url || entry?.item?.['@id'] || entry?.url);
            }
          }
          if (value['@graph']) collectMetadata(value['@graph']);
        }
      };
      collectMetadata(data);
    } catch {
      // Некорректная JSON-LD разметка не мешает извлечению статьи.
    }
  });
  for (const candidate of dateCandidates) {
    if (!candidate) continue;
    publishedDate = normalizePublishedDate(candidate);
    if (publishedDate) {
      break;
    }
  }

  const title = (typeof structuredArticle?.headline === 'string' ? structuredArticle.headline.trim() : '')
    || $('title').first().text().trim()
    || $('h1').first().text().trim()
    || 'Текст страницы';
  const links = resolveUrls(pageUrl, [
    ...$('a[href], link[rel="next"][href]').map((_, element) => $(element).attr('href')).get(),
    ...structuredLinks,
  ])
    .filter((url) => url.startsWith('http'));
  const feeds = $('link[rel~="alternate"][href]').filter((_, element) =>
    /(?:rss|atom)\+xml/i.test($(element).attr('type') || ''),
  ).map((_, element) => $(element).attr('href')).get();
  const clientScripts = $('script[src], script[type="module"]').length > 0;

  const articleNodes = $('article');
  const semanticContent = $('[itemprop="articleBody"], [class*="article-body"], [class*="post-content"]').first();
  const hasArticleMarkup = semanticContent.length > 0
    || $('main > article').length === 1
    || articleNodes.length === 1;
  $('script, style, svg, iframe, nav, footer, header, aside, form, button, [hidden], [aria-hidden="true"]').remove();

  const root = semanticContent.length
    ? semanticContent
    : $('main > article').length === 1
      ? $('main > article').first()
      : articleNodes.length === 1
        ? articleNodes.first()
    : $('main').first().length
      ? $('main').first()
      : $('[role="main"]').first().length
        ? $('[role="main"]').first()
        : $('body');
  const images = [];
  root.find('img').each((_, element) => {
    const image = $(element);
    const pictureSources = image.parent('picture').children('source').toArray().flatMap((source) => {
      const element = $(source);
      return [
        ...responsiveImages(element.attr('data-srcset') || element.attr('srcset')),
        element.attr('data-src'),
      ];
    });
    const urls = resolveUrls(pageUrl, [
      image.attr('data-src'),
      image.attr('data-lazy-src'),
      image.attr('data-original'),
      ...pictureSources,
      ...responsiveImages(image.attr('data-srcset') || image.attr('data-lazy-srcset') || image.attr('srcset')),
      image.attr('src'),
    ]);
    if (urls.length) images.push({ urls, alt: image.attr('alt')?.trim() || '' });
  });
  const hasBodyImages = images.length > 0;
  const coverUrls = resolveUrls(pageUrl, [
    $('meta[property="og:image:secure_url"]').first().attr('content'),
    $('meta[property="og:image"]').first().attr('content'),
    $('meta[name="twitter:image"]').first().attr('content'),
    $('link[rel="image_src"]').first().attr('href'),
    ...[structuredArticle?.image].flat().map((image) => typeof image === 'string' ? image : image?.url),
  ]);
  if (coverUrls.length) images.unshift({
    urls: coverUrls,
    alt: $('meta[property="og:image:alt"]').first().attr('content') || '',
  });
  root.find('noscript').remove();
  const blocks = [];
  root.find('h1, h2, h3, h4, h5, h6, p, li, blockquote, pre, dt, dd, tr').each((_, element) => {
    const tag = element.tagName.toLowerCase();
    const text = tag === 'tr'
      ? $(element).children('th, td').map((__, cell) => $(cell).text().replace(/\s+/g, ' ').trim()).get().join(' | ')
      : $(element).text().replace(/\s+/g, ' ').trim();
    if (text && (blocks.at(-1)?.text !== text)) {
      blocks.push({ text: cleanXmlText(text), tag: tag === 'tr' ? 'p' : tag });
    }
  });

  if (!blocks.length) {
    const text = root.text().replace(/\s+/g, ' ').trim();
    if (text) blocks.push({ text: cleanXmlText(text), tag: 'p' });
  }
  const articleBody = structuredArticle?.articleBody;
  if (typeof articleBody === 'string' && articleBody.trim().length > blocks.reduce((size, block) => size + block.text.length, 0)) {
    blocks.length = 0;
    for (const text of articleBody.split(/\n+/).map((line) => line.trim()).filter(Boolean)) {
      blocks.push({ text: cleanXmlText(text), tag: 'p' });
    }
  }
  const h1Count = root.find('h1').length;
  const h2Count = root.find('h2').length;
  const textLength = blocks.reduce((total, block) => total + block.text.length, 0);
  const paragraphCount = blocks.filter((block) => ['p', 'li', 'blockquote', 'pre'].includes(block.tag)).length;
  const explicitArticle = Boolean(structuredArticle
    || $('meta[property="og:type"]').first().attr('content')?.toLowerCase() === 'article'
    || semanticContent.length
    || (articleNodes.length === 1 && h1Count === 1));
  const isArticle = (explicitArticle && textLength >= 80 && paragraphCount > 0)
    || (hasArticleMarkup && hasBodyImages && (paragraphCount > 0 || h1Count > 0))
    || (textLength >= 350 && paragraphCount >= 3
      && (hasArticleMarkup || (h1Count === 1 && h2Count <= 12 && textLength >= 1000)));

  return {
    title: cleanXmlText(title.replace(/\s+/g, ' ').trim()),
    blocks,
    links,
    feeds,
    clientScripts,
    explicitArticle,
    images: uniqueImages(images),
    isArticle,
    textLength,
    publishedDate,
  };
}

function responsiveImages(srcset = '') {
  const entries = [];
  const urls = /[^,\s][^\s]*/g;
  let match;
  while ((match = urls.exec(srcset))) {
    const url = match[0].replace(/,+$/, '');
    let descriptor = '';
    if (url === match[0]) {
      const comma = srcset.indexOf(',', urls.lastIndex);
      const end = comma === -1 ? srcset.length : comma;
      descriptor = srcset.slice(urls.lastIndex, end).trim();
      urls.lastIndex = end + 1;
    }
    const value = Number.parseFloat(descriptor);
    const width = descriptor.endsWith('x') ? value * 700 : value;
    entries.push({ url, distance: Math.abs((width || 1000) - 1000) });
  }
  return entries
    .sort((a, b) => a.distance - b.distance)
    .map(({ url }) => url);
}

function resolveUrls(pageUrl, candidates) {
  const urls = new Set();
  for (const candidate of candidates) {
    if (typeof candidate !== 'string' || !candidate.trim()) continue;
    try {
      const url = new URL(candidate.trim(), pageUrl);
      if (['http:', 'https:'].includes(url.protocol)) {
        url.hash = '';
        urls.add(url.href);
      } else if (url.protocol === 'data:' && /^data:image\//i.test(url.href)) {
        urls.add(url.href);
      }
    } catch {
      // Некорректный адрес не мешает попробовать следующий вариант.
    }
  }
  return [...urls];
}

function uniqueImages(images) {
  const unique = new Map();
  for (const image of images) {
    const key = image.urls[0];
    const previous = unique.get(key);
    if (previous) {
      previous.urls = [...new Set([...previous.urls, ...image.urls])];
      previous.alt ||= image.alt;
    } else {
      unique.set(key, image);
    }
  }
  return [...unique.values()];
}

async function readImageBytes(response) {
  const declaredSize = Number(response.headers.get('content-length') || 0);
  if (declaredSize > MAX_SOURCE_IMAGE_BYTES) {
    await response.body?.cancel();
    return null;
  }
  const reader = response.body?.getReader();
  if (!reader) return null;
  const chunks = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_SOURCE_IMAGE_BYTES) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }
  return Buffer.concat(chunks);
}

async function fetchImageBytes(address, pageUrl) {
  if (address.startsWith('data:')) {
    const separator = address.indexOf(',');
    if (separator === -1 || address.length - separator > MAX_SOURCE_IMAGE_BYTES * 3) return null;
    const header = address.slice(0, separator);
    const content = address.slice(separator + 1);
    const data = /;base64$/i.test(header)
      ? Buffer.from(content, 'base64')
      : Buffer.from(decodeURIComponent(content));
    return data.length <= MAX_SOURCE_IMAGE_BYTES ? data : null;
  }

  let url = await validatePublicUrl(address);
  const page = new URL(pageUrl);
  const signal = AbortSignal.timeout(8_000);
  for (let redirectCount = 0; redirectCount <= 3; redirectCount += 1) {
    const response = await fetch(url, {
      redirect: 'manual',
      signal,
      headers: {
        'User-Agent': 'SiteToWord/1.0 (+text extraction)',
        Accept: 'image/avif,image/webp,image/svg+xml,image/*,*/*;q=0.5',
        Referer: page.origin === url.origin ? page.href : page.origin + '/',
      },
    });
    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get('location');
      await response.body?.cancel();
      if (!location || redirectCount === 3) return null;
      url = await validatePublicUrl(new URL(location, url).toString());
      continue;
    }
    const contentType = (response.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
    if (!response.ok || (contentType && !contentType.startsWith('image/') && contentType !== 'application/octet-stream')) {
      await response.body?.cancel();
      return null;
    }
    return readImageBytes(response);
  }
  return null;
}

async function downloadImage(reference, pageUrl) {
  for (const url of reference.urls) {
    try {
      const sourceData = await fetchImageBytes(url, pageUrl);
      if (!sourceData?.length) continue;
      const source = sharp(sourceData, { limitInputPixels: 40_000_000 });
      const metadata = await source.metadata();
      if (!metadata.width || !metadata.height) continue;
      if (metadata.format === 'gif') {
        // Проверяем, что GIF декодируется, и сохраняем анимацию в исходном файле.
        await source.stats();
        const scale = Math.min(560 / metadata.width, 680 / metadata.height, 1);
        return {
          url, data: sourceData, type: 'gif',
          width: Math.max(1, Math.round(metadata.width * scale)),
          height: Math.max(1, Math.round(metadata.height * scale)),
          alt: reference.alt,
        };
      }
      const resized = source
        .rotate()
        .resize({ width: 1400, height: 1600, fit: 'inside', withoutEnlargement: true });
      const type = metadata.hasAlpha ? 'png' : 'jpg';
      const { data, info } = await (type === 'png' ? resized.png() : resized.jpeg({ quality: 82 }))
        .toBuffer({ resolveWithObject: true });
      const scale = Math.min(560 / info.width, 680 / info.height, 1);
      return {
        url, data, type,
        width: Math.max(1, Math.round(info.width * scale)),
        height: Math.max(1, Math.round(info.height * scale)),
        alt: reference.alt,
      };
    } catch {
      // Пробуем другой адрес того же изображения.
    }
  }
  return null;
}

async function downloadImages(references, pageUrl, onProgress) {
  const images = [];
  const seen = new Set();
  let totalBytes = 0;
  let skipped = 0;
  for (let offset = 0; offset < references.length; offset += IMAGE_CONCURRENCY) {
    if (totalBytes >= MAX_ARTICLE_IMAGE_BYTES) {
      skipped += references.length - offset;
      onProgress?.(references.length - offset);
      break;
    }
    const batch = references.slice(offset, offset + IMAGE_CONCURRENCY);
    const downloaded = await Promise.all(batch.map((reference) => downloadImage(reference, pageUrl)));
    for (const image of downloaded) {
      if (image && seen.has(image.url)) continue;
      if (!image || totalBytes + image.data.length > MAX_ARTICLE_IMAGE_BYTES) {
        skipped += 1;
        continue;
      }
      seen.add(image.url);
      totalBytes += image.data.length;
      images.push(image);
    }
    onProgress?.(batch.length);
  }
  return { images, skipped };
}

function safeFilename(title) {
  const slug = title
    .normalize('NFKC')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
  return `${cleanXmlText(slug) || 'tekst-stranitsy'}.docx`;
}

async function createWordDocument(articles) {
  const paragraphs = [];
  for (const [index, article] of articles.entries()) {
    const source = `Источник: ${article.url}${article.publishedDate ? ` · Дата публикации: ${article.publishedDate}` : ''}`;
    paragraphs.push(new Paragraph({
      text: cleanXmlText(article.title) || 'Текст страницы',
      heading: HeadingLevel.TITLE,
      ...(index > 0 ? { pageBreakBefore: true } : {}),
    }));
    paragraphs.push(new Paragraph({
      children: [new TextRun({
        text: cleanXmlText(source),
        italics: true,
        color: '666666',
      })],
      spacing: { after: 280 },
    }));

    for (const block of article.blocks) {
      const heading = /^h[1-6]$/.test(block.tag);
      paragraphs.push(new Paragraph({
        text: cleanXmlText(block.text),
        ...(heading
          ? { heading: block.tag === 'h1' ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2 }
          : block.tag === 'li' ? { bullet: { indent: 360 } } : {}),
        spacing: { after: heading ? 140 : 180 },
      }));
    }

    if (article.images?.length) {
      paragraphs.push(new Paragraph({ text: 'Изображения', heading: HeadingLevel.HEADING_2 }));
      for (const image of article.images) {
        paragraphs.push(new Paragraph({
          children: [new ImageRun({
            data: image.data,
            type: image.type,
            transformation: { width: image.width, height: image.height },
          })],
          spacing: { after: 100 },
        }));
        const caption = cleanXmlText(image.alt ?? '');
        if (caption) {
          paragraphs.push(new Paragraph({
            children: [new TextRun({ text: caption, italics: true, color: '666666' })],
            spacing: { after: 180 },
          }));
        }
      }
    }
  }

  const document = new Document({
    creator: 'Site to Word',
    title: articles.length === 1 ? cleanXmlText(articles[0].title) || 'Текст страницы' : `Сборник статей (${articles.length})`,
    sections: [{ properties: {}, children: paragraphs }],
  });
  return Packer.toBuffer(document);
}

function isCrawlableUrl(value, origin) {
  try {
    const url = new URL(value);
    if (url.origin !== origin) return false;
    if (/\.(?:pdf|zip|rar|7z|jpe?g|png|gif|webp|svg|mp[34]|mov|woff2?|ttf|css|js)(?:$|[?#])/i.test(url.pathname)) return false;
    if (/\/(?:login|logout|register|signup|account|profile|search|checkout)(?:\/|$)/i.test(url.pathname)) return false;
    return true;
  } catch {
    return false;
  }
}

function wait(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function robotsDelay(robots) {
  return Math.max(0, Math.min(2_147_483_647, (Number(robots.getCrawlDelay('SiteToWord')) || 0) * 1000));
}

function recordArticleFailure(job, title, error) {
  if (!job) return;
  job.articleFailures += 1;
  job.failedArticles.push({
    title: cleanXmlText(title || 'Без названия'),
    error: cleanXmlText(error?.message || String(error)),
  });
}

function recordRestrictedArticle(job, title) {
  if (!job) return;
  job.restrictedArticles += 1;
  job.restrictedArticleTitles.push(cleanXmlText(title || 'Без названия'));
}

function isKultprofitaMaterialsListing(url) {
  const host = url.hostname.toLowerCase().replace(/^www\./, '');
  return host === 'kultprofita.ru' && url.pathname.replace(/\/+$/, '') === '/lenta/materials';
}

async function loadKultprofitaCatalog(limit, dateFrom, dateTo, job) {
  const apiOrigin = 'https://api.kultprofita.ru';
  const apiRobots = await readRobots(apiOrigin);
  const apiDelay = robotsDelay(apiRobots);
  const pageSize = dateFrom || dateTo ? 100 : Math.min(100, Math.max(20, limit));
  let offset = 0;
  let total = Infinity;
  const entries = [];
  let undatedArticles = 0;
  if (job) {
    job.phase = 'loading-catalog';
    await persistJob(job);
  }

  while (offset < total && entries.length < limit) {
    const endpoint = new URL('/materials/public', apiOrigin);
    endpoint.search = new URLSearchParams({
      limit: String(pageSize),
      offset: String(offset),
      sort: 'new',
      order: 'desc',
    }).toString();
    if (apiRobots.isAllowed(endpoint.toString(), 'SiteToWord') === false) {
      throw new Error('Правила robots.txt запрещают чтение списка материалов');
    }
    if (offset) await wait(Math.max(250, apiDelay));
    const page = await fetchJsonPage(endpoint, apiOrigin, apiRobots);
    if (!Array.isArray(page.items)) throw new Error('Не удалось прочитать список материалов сайта');
    total = Number(page.total || 0);
    if (job) {
      job.catalogTotal = total;
      job.catalogRead = offset;
      job.phase = 'loading-catalog';
    }

    for (const item of page.items) {
      const publishedDate = normalizePublishedDate(item.published_at);
      if ((!publishedDate && (dateFrom || dateTo))) {
        undatedArticles += 1;
        continue;
      }
      if ((dateFrom && publishedDate < dateFrom) || (dateTo && publishedDate > dateTo)) continue;
      if (item.can_read === false || item.effective_access_type !== 'public') {
        recordRestrictedArticle(job, item.title);
        continue;
      }
      const slug = item.published_slug || item.slug;
      if (!slug) continue;
      entries.push({
        title: item.title || item.seo_title || slug,
        slug,
        publishedDate,
        coverAssetId: item.cover_asset_id,
      });
      if (job) job.articlesFound = entries.length;
      if (entries.length >= limit) break;
    }

    offset += page.items.length;
    if (job) job.catalogRead = Math.min(offset, total);
    if (!page.items.length) break;
    if (job) await persistJob(job);
  }

  const selected = entries.slice(0, Number.isFinite(limit) ? limit : entries.length);
  if (job) {
    job.articlesFound = selected.length;
    job.articlesTotal = selected.length;
    job.articlesProcessed = 0;
    job.undatedArticles = undatedArticles;
    job.phase = 'loading-articles';
    await persistJob(job);
  }
  return { entries: selected, apiOrigin, apiRobots, apiDelay, undatedArticles };
}

async function fetchKultprofitaArticle(entry, apiOrigin, apiRobots) {
  const endpoint = new URL(`/materials/public/${encodeURIComponent(entry.slug)}`, apiOrigin);
  if (apiRobots.isAllowed(endpoint.toString(), 'SiteToWord') === false) {
    throw new Error('Правила robots.txt запрещают чтение материала');
  }
  const item = await fetchJsonPage(endpoint, apiOrigin, apiRobots);
  if (item.can_read === false || !item.published_content_json) {
    throw new Error('У этого материала нет публичного текста');
  }

  const blocks = item.published_content_json.blocks || [];
  const html = blocks.map((block) => {
    if (block?.data?.html) return block.data.html;
    const text = block?.data?.text;
    return typeof text === 'string' && text.trim()
      ? '<p>' + text.replaceAll('&', '&amp;').replaceAll('<', '&lt;') + '</p>'
      : '';
  }).join('\n');
  if (!html.trim()) throw new Error('Не удалось извлечь текст материала');

  const slug = item.published_slug || item.slug || entry.slug;
  const articleUrl = `https://kultprofita.ru/materials/${encodeURIComponent(slug)}`;
  const extracted = extractPage(html, `${apiOrigin}/`);
  if (!extracted.blocks.length) throw new Error('Не удалось извлечь текст материала');
  const images = [...extracted.images];
  const assetUrl = (id) => `${apiOrigin}/files/public/assets/${encodeURIComponent(id)}`;
  for (const block of blocks) {
    const assetIds = block?.data?.asset_ids;
    if (!Array.isArray(assetIds)) continue;
    for (const id of assetIds) {
      if (id) images.push({ urls: [assetUrl(id)], alt: '' });
    }
  }
  const coverAssetId = item.cover_asset_id || entry.coverAssetId || item.seo_image_asset_id;
  if (coverAssetId) {
    images.unshift({ urls: [assetUrl(coverAssetId)], alt: item.title || entry.title });
  }
  return {
    title: item.title || entry.title,
    blocks: extracted.blocks,
    images: uniqueImages(images),
    url: articleUrl,
    publishedDate: normalizePublishedDate(item.published_at) || entry.publishedDate,
    textLength: extracted.textLength,
  };
}

async function crawlKultprofitaArticles(catalog, job) {
  const { entries } = catalog;
  const results = new Array(entries.length);
  let nextIndex = 0;
  let totalText = 0;
  let fatalError = null;

  job.phase = 'loading-articles';
  job.pagesVisited = 0;
  job.pagesQueued = entries.length;
  job.articlesFound = entries.length;
  job.articlesTotal = entries.length;
  job.articlesProcessed = 0;

  const worker = async () => {
    while (!fatalError) {
      const index = nextIndex;
      nextIndex += 1;
      if (index >= entries.length) return;

      const entry = entries[index];
      job.pagesVisited = Math.min(entries.length, job.pagesVisited + 1);
      job.pagesQueued = Math.max(0, entries.length - job.articlesProcessed);
      job.currentArticle = entry.title;
      if (catalog.apiDelay) await wait(catalog.apiDelay);

      try {
        const article = await fetchKultprofitaArticle(entry, catalog.apiOrigin, catalog.apiRobots);
        if (totalText + article.textLength > MAX_TOTAL_TEXT) {
          fatalError = new Error('Общий объём текста превысил допустимый лимит');
          return;
        }

        totalText += article.textLength;
        results[index] = article;
        job.articlesProcessed += 1;
        job.pagesQueued = Math.max(0, entries.length - job.articlesProcessed);
        if (job.articlesProcessed % 10 === 0 || job.articlesProcessed === entries.length) {
          await persistJob(job);
        }
      } catch (error) {
        if (error.message === 'Общий объём текста превысил допустимый лимит') {
          fatalError = error;
          return;
        }
        recordArticleFailure(job, entry.title, error);
        job.pagesQueued = Math.max(0, entries.length - job.articlesProcessed - job.articleFailures);
        await persistJob(job);
      }
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(catalog.apiDelay ? 1 : KULTPROFITA_CONCURRENCY, entries.length) }, () => worker()),
  );
  if (fatalError) throw fatalError;

  const articles = results.filter(Boolean);
  if (!articles.length) {
    throw new Error('Не удалось загрузить ни одной статьи из списка материалов');
  }
  job.articlesProcessed = articles.length;
  job.articlesTotal = articles.length;
  job.undatedArticles = catalog.undatedArticles;
  job.crawlComplete = articles.length === entries.length;
  job.phase = 'preparing-documents';
  return {
    articles,
    undatedArticles: catalog.undatedArticles,
    complete: job.crawlComplete,
  };
}

function identifyPublicBlog(url) {
  const host = url.hostname.replace(/^www\./, '');
  const parts = url.pathname.split('/').filter(Boolean);
  if (parts.shift() !== 'blog') return null;
  try {
    for (const [index, part] of parts.entries()) parts[index] = decodeURIComponent(part);
  } catch {
    return null;
  }
  if (/^one-vv\d+\.com$/.test(host) && parts.length <= 2) {
    const [category, slug] = parts;
    return { kind: 'one', category, slug };
  }
  if (host === 'csgo500.com') {
    if (parts.length <= 1) return { kind: 'csg', slug: parts[0] };
    if (parts.length === 2 && parts[0] === 'category') {
      return { kind: 'csg', category: parts[1] };
    }
  }
  return null;
}

function blogArticle(data, pageUrl, kind) {
  const one = kind === 'one';
  const html = one ? data.text : data.html;
  if (typeof html !== 'string' || !html.trim()) return null;
  const content = one ? html : html.replaceAll('cdn://', 'https://cdnv1.csgo500.io/');
  const extracted = extractPage(content, pageUrl);
  if (!extracted.blocks.length) return null;
  const title = cleanXmlText((one ? data.header : data.title) || extracted.title);
  const cover = one ? data.picture : String(data.bannerImage || '').replace('cdn://', 'https://cdnv1.csgo500.io/');
  const urls = resolveUrls(pageUrl, [cover]);
  const images = urls.length ? [{ urls, alt: title }, ...extracted.images] : extracted.images;
  return {
    title,
    blocks: extracted.blocks,
    images: uniqueImages(images),
    textLength: extracted.textLength,
    publishedDate: normalizePublishedDate(one ? data.publishedAt : data.createDate) || extracted.publishedDate,
    url: pageUrl,
  };
}

async function crawlPublicBlog(startUrl, source, robots, limit, dateFrom, dateTo, job) {
  const origin = startUrl.origin;
  const initialPage = await fetchPage(startUrl, origin, robots);
  const one = source.kind === 'one';
  const document = load(initialPage.html);
  const locale = document('html').attr('lang') || 'ru-RU';
  const delay = robotsDelay(robots);
  const concurrency = delay ? 1 : PAGE_CONCURRENCY;
  const articles = [];
  const seen = new Set();
  const cursors = new Set();
  let totalText = 0;
  let undatedArticles = 0;

  const readApi = async (path, parameters) => {
    if (delay) await wait(delay);
    const endpoint = new URL(path, origin);
    endpoint.search = new URLSearchParams(parameters);
    return fetchJsonPage(endpoint, origin, robots);
  };
  const readOne = (categorySlug, slug) =>
    readApi('/api/blog/v3/public/posts/get/by/slug', { locale, categorySlug, slug });
  const readCsg = async (slug) => (await readApi('/api/blog/' + encodeURIComponent(slug))).blog;
  const articleUrl = (category, slug) => new URL(
    '/blog/' + (one ? encodeURIComponent(category) + '/' : '') + encodeURIComponent(slug), origin,
  ).href;
  const eligible = (url) => robots.isAllowed(url, CRAWLER_NAME) !== false;
  const addArticle = (data, url) => {
    if (!data || !eligible(url)) return;
    const article = blogArticle(data, url, source.kind);
    if (!article) return;
    const { publishedDate } = article;
    if (!publishedDate && (dateFrom || dateTo)) {
      undatedArticles += 1;
      return;
    }
    if ((dateFrom && publishedDate < dateFrom) || (dateTo && publishedDate > dateTo)) return;
    totalText += article.textLength;
    if (totalText > MAX_TOTAL_TEXT) throw new Error('Общий объём текста превысил допустимый лимит');
    articles.push(article);
    job.articlesFound = articles.length;
    job.articlesProcessed = articles.length;
  };

  job.phase = 'scanning';
  if (source.slug) {
    job.pagesVisited = 1;
    const data = one ? await readOne(source.category, source.slug) : await readCsg(source.slug);
    addArticle(data, articleUrl(source.category, source.slug));
  } else {
    let cursor = one ? '' : 1;
    while (cursor !== null && !cursors.has(cursor) && cursors.size < MAX_CRAWL_PAGES
      && seen.size < MAX_CRAWL_PAGES && articles.length < limit) {
      cursors.add(cursor);
      const parameters = one
        ? { locale, limit: String(Math.min(40, Math.max(10, limit - articles.length))) }
        : { page: String(cursor) };
      if (source.category) parameters[one ? 'categorySlug' : 'category'] = source.category;
      if (one && cursor) parameters.cursor = cursor;
      const page = await readApi(one ? '/api/blog/v3/public/posts/get/list' : '/api/blogs', parameters);
      if (!Array.isArray(one ? page.items : page.blogs)) {
        throw new Error('API сайта не вернул список статей');
      }
      const active = Array.isArray(page.otherActiveBlogs) ? page.otherActiveBlogs : [];
      const items = one ? page.items : [...page.blogs, ...active];
      cursor = one ? page.meta?.cursor || null : page.pagination?.nextPage ? cursor + 1 : null;
      if (!items.length) cursor = null;
      const entries = [];
      for (const item of items) {
        const slug = one ? item.slug : item.niceUrl || item._id;
        if (!slug || (one && !item.categorySlug)) continue;
        const url = articleUrl(item.categorySlug, slug);
        if (seen.has(url) || !eligible(url)) continue;
        seen.add(url);
        entries.push({ item, slug, url });
        if (seen.size >= MAX_CRAWL_PAGES) break;
      }
      job.pagesQueued = seen.size;
      for (let next = 0; next < entries.length && articles.length < limit;) {
        const batch = entries.slice(next, next + Math.min(concurrency, limit - articles.length));
        next += batch.length;
        const results = await Promise.allSettled(batch.map(({ item, slug }) =>
          one ? Promise.resolve(item) : readCsg(slug),
        ));
        for (const [index, result] of results.entries()) {
          job.pagesVisited += 1;
          if (result.status === 'fulfilled') {
            addArticle(result.value, batch[index].url);
          } else {
            recordArticleFailure(job, batch[index].item.title, result.reason);
          }
        }
      }
    }
    if (!articles.length && job.failedArticles.length) throw new Error(job.failedArticles[0].error);
    return { articles, undatedArticles, complete: articles.length >= limit || cursor === null };
  }
  return { articles, undatedArticles, complete: true };
}

function parseDiscoveryXml(xml, baseUrl) {
  const $ = load(xml, { xml: true });
  const root = $.root().children().first();
  const name = (node) => node.name?.split(':').at(-1);
  const children = (node, tag) => $(node).children().filter((_, child) => name(child) === tag);
  const get = (node, tag) => children(node, tag).first();
  const rootName = name(root.get(0) || {});
  if (rootName === 'sitemapindex' || rootName === 'urlset') {
    const index = rootName === 'sitemapindex';
    const urls = resolveUrls(baseUrl, children(root, index ? 'sitemap' : 'url')
      .map((_, item) => get(item, 'loc').text().trim()).get());
    return { sitemaps: index ? urls : [], pages: index ? [] : urls, articles: [] };
  }
  const atom = rootName === 'feed';
  if (!atom && rootName !== 'rss') return { sitemaps: [], pages: [], articles: [] };
  const container = atom ? root : get(root, 'channel');
  const pages = [];
  const articles = [];
  children(container, atom ? 'entry' : 'item').each((_, item) => {
    const link = children(item, 'link').filter((_, child) => !$(child).attr('rel') || $(child).attr('rel') === 'alternate').first();
    const [url] = resolveUrls(baseUrl, [atom ? link.attr('href') : link.text()]);
    if (!url) return;
    pages.push(url);
    const content = get(item, atom ? 'content' : 'encoded');
    const html = content.attr('type') === 'xhtml' ? content.html() : content.text();
    if (!html?.trim() || content.attr('src')) return;
    const extracted = extractPage(html, url);
    const title = cleanXmlText(get(item, 'title').text().trim()) || extracted.title;
    articles.push({
      ...extracted,
      url,
      title,
      publishedDate: normalizePublishedDate(get(item, atom ? 'published' : 'pubDate').text().trim()) || extracted.publishedDate,
    });
  });
  return { sitemaps: [], pages, articles };
}

function inSitemapSection(url, startUrl) {
  const scope = startUrl.pathname.replace(/\/+$/, '');
  if (!scope) return true;
  const path = new URL(url).pathname.replace(/\/+$/, '');
  return path === scope || path.startsWith(scope + '/');
}

async function discoverAdditionalPages(startUrl, robots, feedUrls) {
  const origin = startUrl.origin;
  const section = startUrl.pathname.split('/').filter(Boolean).at(-1)?.toLowerCase();
  const rank = (url) => section && new URL(url, origin).pathname.toLowerCase().includes(section) ? 0 : 1;
  const advertised = resolveUrls(origin, robots.getSitemaps()).sort((a, b) => rank(a) - rank(b));
  const queue = [...feedUrls, ...advertised.filter((url) => !rank(url)), new URL('/sitemap.xml', origin).href, ...advertised];
  const visited = new Set();
  const pages = new Set();
  const articles = [];
  const delay = robotsDelay(robots);
  for (let index = 0; index < queue.length && visited.size < MAX_DISCOVERY_DOCUMENTS; index += 1) {
    const url = queue[index];
    if (!isCrawlableUrl(url, origin) || visited.has(url) || robots.isAllowed(url, CRAWLER_NAME) === false) continue;
    visited.add(url);
    if (delay) await wait(delay);
    try {
      const document = await fetchPage(new URL(url), origin, robots, true);
      if (!document.isXml) continue;
      const found = parseDiscoveryXml(document.html, document.url);
      if (found.sitemaps.length) {
        queue.splice(index + 1, 0, ...found.sitemaps.sort((a, b) => rank(a) - rank(b)).slice(0, MAX_DISCOVERY_DOCUMENTS));
      }
      const postSitemap = /\/posts?[-_.]/i.test(new URL(document.url).pathname);
      for (const page of found.pages) {
        if (!isCrawlableUrl(page, origin)) continue;
        if (feedUrls.has(url) || postSitemap || inSitemapSection(page, startUrl)) pages.add(page);
      }
      articles.push(...found.articles);
    } catch {
      // Карта сайта и ленты дополняют обычный обход.
    }
  }
  const articlePath = /\/(?:blog|posts?|news|articles?|journal|stories)\//i;
  return {
    pages: [...pages].sort((a, b) => Number(articlePath.test(b)) - Number(articlePath.test(a))),
    articles,
  };
}

async function crawlArticles(startUrl, limit, dateFrom, dateTo, job) {
  const origin = startUrl.origin;
  job.phase = 'checking-rules';
  const robots = await readRobots(origin);
  checkRobots(startUrl, robots);
  if (isKultprofitaMaterialsListing(startUrl)) {
    const catalog = await loadKultprofitaCatalog(limit, dateFrom, dateTo, job);
    return crawlKultprofitaArticles(catalog, job);
  }
  const publicBlog = identifyPublicBlog(startUrl);
  if (publicBlog) {
    const result = await crawlPublicBlog(startUrl, publicBlog, robots, limit, dateFrom, dateTo, job);
    if (!result.articles.length) throw new Error('В публичном блоге не найдены статьи за выбранный период');
    return result;
  }

  const delay = robotsDelay(robots);
  const concurrency = delay ? 1 : PAGE_CONCURRENCY;
  const maxVisited = limit === Infinity ? MAX_CRAWL_PAGES : Math.min(5000, Math.max(100, limit * 4));
  const maxDepth = limit === Infinity ? Infinity : 3;
  const queue = [{ url: startUrl.href, depth: 0 }];
  const discovered = new Set([startUrl.href]);
  const articleUrls = new Set();
  const feedUrls = new Set();
  const articles = [];
  let next = 0;
  let pagesVisited = 0;
  let totalText = 0;
  const undatedUrls = new Set();
  let usedDiscovery = false;
  let sawClientShell = false;
  job.phase = 'scanning';

  const queueLink = (url, depth, fromSitemap = false) => {
    if (discovered.has(url) || !isCrawlableUrl(url, origin) || robots.isAllowed(url, CRAWLER_NAME) === false) return;
    discovered.add(url);
    queue.push({ url, depth, fromSitemap });
  };
  const addArticle = (page) => {
    if (articles.length >= limit || articleUrls.has(page.url) || !isCrawlableUrl(page.url, origin)
      || robots.isAllowed(page.url, CRAWLER_NAME) === false || !page.blocks.length) return;
    const { publishedDate } = page;
    if (!publishedDate && (dateFrom || dateTo)) {
      undatedUrls.add(page.url);
      return;
    }
    undatedUrls.delete(page.url);
    if ((dateFrom && publishedDate < dateFrom) || (dateTo && publishedDate > dateTo)) return;
    totalText += page.textLength;
    if (totalText > MAX_TOTAL_TEXT) throw new Error('Общий объём текста превысил допустимый лимит');
    articleUrls.add(page.url);
    const { title, blocks, url, images } = page;
    articles.push({ title, blocks, url, publishedDate, images });
  };

  while (pagesVisited < maxVisited && articles.length < limit) {
    if (next === queue.length) {
      if (usedDiscovery) break;
      usedDiscovery = true;
      const extra = await discoverAdditionalPages(startUrl, robots, feedUrls);
      for (const article of extra.articles) addArticle(article);
      for (const url of extra.pages) queueLink(url, 1, true);
      if (next === queue.length || articles.length >= limit) break;
    }
    const count = Math.min(concurrency, maxVisited - pagesVisited);
    const batch = queue.slice(next, next + count);
    next += batch.length;
    if (delay && pagesVisited) await wait(delay);
    const pages = await Promise.all(batch.map(async ({ url, depth }) => {
      try {
        return await fetchPage(new URL(url), origin, robots, true);
      } catch (error) {
        if (!depth) throw error;
        return null;
      }
    }));
    pagesVisited += batch.length;

    for (const [index, response] of pages.entries()) {
      if (!response || articles.length >= limit) continue;
      const entry = batch[index];
      if (response.isXml) {
        const found = parseDiscoveryXml(response.html, response.url);
        for (const article of found.articles) addArticle(article);
        for (const url of found.pages) queueLink(url, entry.depth + 1, true);
        for (const url of found.sitemaps) feedUrls.add(url);
        continue;
      }
      const page = { ...extractPage(response.html, response.url), url: response.url };
      const initial = entry.depth === 0;
      if (initial) {
        sawClientShell = page.clientScripts && !page.links.length && page.textLength < 100;
        for (const href of resolveUrls(page.url, page.feeds)) feedUrls.add(href);
      }
      if ((page.isArticle && (!initial || page.explicitArticle))
        || (entry.fromSitemap && page.textLength >= 160 && page.blocks.filter(({ tag }) => tag === 'p').length >= 2)) {
        addArticle(page);
      }
      const depth = entry.depth + 1;
      if (depth > maxDepth) continue;
      for (const url of page.links) queueLink(url, depth);
    }
    job.pagesVisited = pagesVisited;
    job.pagesQueued = queue.length;
    job.articlesFound = articles.length;
    job.articlesProcessed = articles.length;
  }

  if (!articles.length) {
    if (sawClientShell) {
      throw new Error('Сайт загружает статьи через JavaScript. Публичная лента, карта сайта и HTML не содержат доступного текста статей.');
    }
    if (dateFrom || dateTo) throw new Error('Не найдено статей за выбранный период');
    throw new Error('Статьи не найдены. Укажите ссылку на публичную статью, раздел блога или RSS-ленту.');
  }
  job.articlesFound = articles.length;
  job.articlesProcessed = articles.length;
  const complete = limit !== Infinity || (next === queue.length && pagesVisited < maxVisited);
  return { articles, undatedArticles: undatedUrls.size, complete };
}

function validateDate(value, label) {
  if (!value) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(`Некорректная дата «${label}»`);
  }
  const parsed = new Date(`${value}T00:00:00.000Z`);
  if (Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== value) {
    throw new Error(`Некорректная дата «${label}»`);
  }
  return value;
}

async function* readArchiveEntry(path) {
  yield* createReadStream(path);
}

async function createArticleArchive(articles, siteName, job) {
  const zip = new JSZip();
  const folder = zip.folder(siteName);
  const documentsDirectory = join(jobDataDirectory, job.id, 'documents');
  await mkdir(documentsDirectory, { recursive: true });
  job.phase = 'preparing-documents';
  job.articlesTotal = articles.length;
  job.imagesToProcess = articles.reduce((count, article) => count + article.images.length, 0);
  await persistJob(job);

  const imageFailures = new Map();
  let documentBytes = 0;
  let archiveTooLarge = false;
  let nextDocument = 0;
  const prepareWorker = async () => {
    while (!archiveTooLarge) {
      const index = nextDocument;
      nextDocument += 1;
      if (index >= articles.length) return;
      const article = articles[index];
      try {
        const { images, skipped } = await downloadImages(article.images, article.url, (count) => {
          job.imagesProcessed += count;
        });
        const document = await createWordDocument([{ ...article, images }]);
        await assertReadableDocx(document);
        // JSZip не записывает ZIP64, поэтому оставляем запас для заголовков архива.
        if (documentBytes + document.length > MAX_ARCHIVE_DOCUMENT_BYTES) {
          archiveTooLarge = true;
          return;
        }
        documentBytes += document.length;
        const number = String(index + 1).padStart(3, '0');
        const filename = `${number}_${safeFilename(article.title)}`;
        const documentPath = join(documentsDirectory, filename);
        await writeFile(documentPath, document);
        folder.file(filename, Readable.from(readArchiveEntry(documentPath), { objectMode: false }));
        job.preparedArticles += 1;
        job.imagesIncluded += images.length;
        job.imagesFound = job.imagesIncluded;
        job.imagesSkipped += skipped;
        if (skipped) imageFailures.set(index, { filename, skipped });
      } catch (error) {
        recordArticleFailure(job, article.title, error);
      } finally {
        article.blocks = [];
        article.images = [];
      }
      if ((job.preparedArticles + job.articleFailures) % 10 === 0) await persistJob(job);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(DOCX_CONCURRENCY, articles.length) }, () => prepareWorker()),
  );
  if (archiveTooLarge) throw new Error('Архив слишком большой. Уменьшите количество статей или задайте период выгрузки.');
  if (!job.preparedArticles) throw new Error('Не удалось собрать ни одного Word-файла');

  folder.file(
    'README.txt',
    'Распакуйте ZIP-архив и открывайте файлы .docx в Microsoft Word или совместимом редакторе. Изображения размещены в конце документа статьи.\n',
  );
  const report = [];
  if (job.restrictedArticleTitles.length) {
    report.push('Материалы, недоступные без соответствующего доступа:');
    report.push(...job.restrictedArticleTitles.map((title) => `- ${title}`), '');
  }
  if (job.failedArticles.length) {
    report.push('Материалы с временной или технической ошибкой:');
    report.push(...job.failedArticles.map(({ title, error }) => `- ${title}: ${error}`), '');
  }
  if (imageFailures.size) {
    report.push('Изображения, которые не удалось скачать или встроить:');
    for (const [, { filename, skipped }] of [...imageFailures.entries()].sort(([a], [b]) => a - b)) {
      report.push(`- ${filename}: ${skipped}`);
    }
  }
  if (report.length) folder.file('Отчет.txt', report.join('\n'));

  job.phase = 'packing-zip';
  job.zipPercent = 0;
  await persistJob(job);
  const archivePath = join(jobDataDirectory, job.id, `${siteName}.zip`);
  await pipeline(
    zip.generateNodeStream(
      {
        type: 'nodebuffer',
        // Записываем размеры и CRC в заголовки для совместимости с распаковщиками.
        streamFiles: false,
        compression: 'STORE',
      },
      (metadata) => {
        job.zipPercent = Math.floor(metadata.percent);
      },
    ),
    createWriteStream(archivePath),
  );
  const archiveStats = await stat(archivePath);
  if (!archiveStats.size) throw new Error('Получился пустой ZIP-архив');
  await rm(documentsDirectory, { recursive: true, force: true });
  job.archivePath = archivePath;
  job.archiveSize = archiveStats.size;
  job.zipPercent = 100;
}

async function assertReadableDocx(buffer) {
  const documentZip = await JSZip.loadAsync(buffer);
  const requiredParts = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'];
  if (requiredParts.some((part) => !documentZip.file(part))) {
    throw new Error('Не удалось собрать корректный Word-файл. Повторите выгрузку');
  }
  for (const part of documentZip.file(/\.(?:xml|rels)$/)) {
    try {
      const xml = await part.async('text');
      new SaxesParser({ xmlns: true }).write(xml).close();
    } catch (error) {
      throw new Error(`Word-файл повреждён при создании (${part.name})`, { cause: error });
    }
  }
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  response.end(JSON.stringify(payload));
}

async function readJson(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_REQUEST_BYTES) throw new Error('Запрос слишком большой');
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

function publicJobStatus(job) {
  return {
    id: job.id,
    status: job.status,
    phase: job.phase,
    limit: job.limit,
    pagesVisited: job.pagesVisited,
    pagesQueued: job.pagesQueued,
    catalogRead: job.catalogRead,
    catalogTotal: job.catalogTotal,
    articlesFound: job.articlesFound,
    articlesProcessed: job.articlesProcessed,
    articleFailures: job.articleFailures,
    articlesTotal: job.articlesTotal,
    preparedArticles: job.preparedArticles,
    imagesFound: job.imagesFound,
    imagesIncluded: job.imagesIncluded,
    imagesSkipped: job.imagesSkipped || 0,
    imagesProcessed: job.imagesProcessed,
    imagesToProcess: job.imagesToProcess,
    zipPercent: job.zipPercent,
    undatedArticles: job.undatedArticles,
    restrictedArticles: job.restrictedArticles,
    crawlComplete: job.crawlComplete,
    error: job.error,
  };
}

function persistJob(job) {
  const snapshot = JSON.stringify({
    ...publicJobStatus(job),
    createdAt: job.createdAt,
    finishedAt: job.finishedAt,
    siteName: job.siteName,
    filename: job.filename,
    archivePath: job.archivePath,
    archiveSize: job.archiveSize,
    failedArticles: job.failedArticles || [],
    restrictedArticleTitles: job.restrictedArticleTitles || [],
    request: job.request || null,
    restartCount: job.restartCount || 0,
  });
  job.persistPromise = (job.persistPromise || Promise.resolve())
    .catch(() => {})
    .then(async () => {
      const temporaryPath = join(jobDataDirectory, `${job.id}.${randomUUID()}.tmp`);
      await writeFile(temporaryPath, snapshot, 'utf8');
      await rename(temporaryPath, join(jobDataDirectory, `${job.id}.json`));
    })
    .catch(() => {});
  return job.persistPromise;
}

async function restoreJobs() {
  await mkdir(jobDataDirectory, { recursive: true });
  const files = await readdir(jobDataDirectory);
  const expiration = Date.now() - 24 * 60 * 60 * 1000;
  for (const filename of files.filter((name) => name.endsWith('.json'))) {
    const metadataPath = join(jobDataDirectory, filename);
    try {
      const stored = JSON.parse(await readFile(metadataPath, 'utf8'));
      const fileId = filename.slice(0, -'.json'.length);
      if (!/^[a-f0-9-]{36}$/i.test(fileId) || stored.id !== fileId) {
        await rm(metadataPath, { force: true });
        continue;
      }
      if (stored.createdAt < expiration) {
        await rm(join(jobDataDirectory, fileId), { recursive: true, force: true });
        await rm(metadataPath, { force: true });
        continue;
      }
      const job = {
        ...stored,
        persistPromise: null,
        articleFailures: Number(stored.articleFailures || 0),
        failedArticles: Array.isArray(stored.failedArticles) ? stored.failedArticles : [],
        restrictedArticles: Number(stored.restrictedArticles || 0),
        restrictedArticleTitles: Array.isArray(stored.restrictedArticleTitles) ? stored.restrictedArticleTitles : [],
      };
      if (['queued', 'running'].includes(job.status)) {
        if (job.request && (job.restartCount || 0) < 1) {
          job.restartCount = (job.restartCount || 0) + 1;
          job.status = 'queued';
          job.phase = 'queued';
          job.error = null;
          jobs.set(job.id, job);
          await persistJob(job);
          const options = {
            ...job.request,
            limit: job.request.limit === 'max' ? Infinity : Number(job.request.limit),
          };
          setImmediate(() => void runCrawlJob(job, options));
          continue;
        }
        job.status = 'error';
        job.phase = 'error';
        job.error = 'Сбор прервался при перезапуске сервера. Запустите его повторно.';
        job.finishedAt = Date.now();
      } else if (job.status === 'done') {
        try {
          if (!job.archivePath) throw new Error('Archive path missing');
          if (!job.archivePath.startsWith(`${jobDataDirectory}${sep}${job.id}${sep}`)) {
            throw new Error('Invalid archive location');
          }
          await stat(job.archivePath);
        } catch {
          job.status = 'error';
          job.phase = 'error';
          job.error = 'Архив больше не найден. Запустите сбор повторно.';
          job.finishedAt = Date.now();
        }
      }
      jobs.set(job.id, job);
      if (job.status === 'error') await persistJob(job);
    } catch {
      await rm(metadataPath, { force: true });
    }
  }
}

async function pruneJobs() {
  const expiration = Date.now() - 24 * 60 * 60 * 1000;
  for (const [id, job] of jobs) {
    if (job.createdAt < expiration && ['done', 'error'].includes(job.status)) {
      jobs.delete(id);
      await rm(join(jobDataDirectory, id), { recursive: true, force: true });
      await rm(join(jobDataDirectory, `${id}.json`), { force: true });
    }
  }
}

async function runCrawlJob(job, options) {
  job.status = 'running';
  job.phase = 'checking-rules';
  job.error = null;
  job.pagesVisited = 0;
  job.pagesQueued = 0;
  job.catalogRead = 0;
  job.catalogTotal = null;
  job.articlesFound = 0;
  job.articlesProcessed = 0;
  job.articleFailures = 0;
  job.failedArticles = [];
  job.restrictedArticles = 0;
  job.restrictedArticleTitles = [];
  job.articlesTotal = null;
  job.preparedArticles = 0;
  job.imagesFound = 0;
  job.imagesIncluded = 0;
  job.imagesSkipped = 0;
  job.imagesProcessed = 0;
  job.imagesToProcess = 0;
  job.zipPercent = 0;
  job.archivePath = null;
  job.archiveSize = 0;
  await rm(join(jobDataDirectory, job.id), { recursive: true, force: true });
  await persistJob(job);
  try {
    const safeUrl = await validatePublicUrl(options.url);
    job.siteName = safeUrl.hostname.replace(/[^a-zA-Z0-9.-]/g, '_');
    const result = await crawlArticles(safeUrl, options.limit, options.dateFrom, options.dateTo, job);
    await createArticleArchive(result.articles, job.siteName, job);
    job.articlesTotal = job.preparedArticles;
    job.undatedArticles = result.undatedArticles;
    job.crawlComplete = result.complete
      && job.preparedArticles === result.articles.length
      && job.articleFailures === 0;
    job.filename = `${job.siteName}.zip`;
    job.status = 'done';
    job.phase = 'complete';
    job.finishedAt = Date.now();
    await persistJob(job);
  } catch (error) {
    job.status = 'error';
    job.phase = 'error';
    job.error = error.name === 'TimeoutError'
      ? 'Сайт отвечает слишком долго. Попробуйте позже'
      : error.message || 'Не удалось обработать сайт';
    job.finishedAt = Date.now();
    await rm(join(jobDataDirectory, job.id), { recursive: true, force: true }).catch(() => {});
    await persistJob(job);
  }
}

const server = createServer(async (request, response) => {
  const pathname = new URL(request.url, 'http://localhost').pathname;
  if (request.method === 'GET' && pathname === '/') {
    try {
      const html = await readFile(join(directory, 'index.html'));
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      response.end(html);
    } catch {
      sendJson(response, 500, { error: 'Не удалось открыть интерфейс' });
    }
    return;
  }

  const jobMatch = pathname.match(/^\/api\/jobs\/([a-f0-9-]+)(?:\/(download))?$/i);
  if (request.method === 'GET' && jobMatch) {
    await pruneJobs();
    const job = jobs.get(jobMatch[1]);
    if (!job) {
      sendJson(response, 404, { error: 'Задача не найдена или уже удалена' });
      return;
    }
    if (jobMatch[2] === 'download') {
      if (job.status !== 'done' || !job.archivePath) {
        sendJson(response, 409, { error: 'Архив ещё не готов' });
        return;
      }
      let archiveStats;
      try {
        archiveStats = await stat(job.archivePath);
      } catch {
        sendJson(response, 410, { error: 'Архив не найден на диске. Запустите сбор повторно' });
        return;
      }
      response.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="${job.filename}"`,
        'Content-Length': job.archiveSize || archiveStats.size,
        'X-Articles-Count': String(job.articlesTotal),
        'X-Images-Count': String(job.imagesFound),
        'X-Images-Skipped': String(job.imagesSkipped || 0),
        'X-Article-Failures': String(job.articleFailures),
        'X-Restricted-Articles': String(job.restrictedArticles),
        'X-Undated-Articles': String(job.undatedArticles),
        'X-Crawl-Complete': String(job.crawlComplete),
        'Cache-Control': 'no-store',
      });
      await pipeline(createReadStream(job.archivePath), response).catch(() => {});
      return;
    }
    sendJson(response, 200, publicJobStatus(job));
    return;
  }

  if (request.method === 'POST' && pathname === '/api/export') {
    try {
      const body = await readJson(request);
      if (typeof body.url !== 'string' || !body.url.trim()) {
        sendJson(response, 400, { error: 'Введите ссылку на сайт' });
        return;
      }
      const safeUrl = await validatePublicUrl(body.url.trim());
      let article;
      if (identifyPublicBlog(safeUrl)?.slug) {
        const result = await crawlArticles(safeUrl, 1, '', '', {});
        [article] = result.articles;
      } else {
        const page = await fetchPage(safeUrl);
        article = { ...extractPage(page.html, page.url), url: page.url };
      }
      const { images } = await downloadImages(article.images, article.url);
      if (!article.blocks.length && !images.length) throw new Error('На странице не найден текст или изображения');
      const document = await createWordDocument([{ ...article, images }]);
      await assertReadableDocx(document);
      response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="page.docx"; filename*=UTF-8''${encodeURIComponent(safeFilename(article.title))}`,
        'Content-Length': document.length,
        'Cache-Control': 'no-store',
      });
      response.end(document);
    } catch (error) {
      const message = error.name === 'TimeoutError'
        ? 'Сайт отвечает слишком долго. Попробуйте позже'
        : error instanceof SyntaxError
          ? 'Некорректный запрос'
          : error.message || 'Не удалось обработать сайт';
      sendJson(response, 400, { error: message });
    }
    return;
  }

  if (request.method === 'POST' && pathname === '/api/crawl') {
    try {
      const body = await readJson(request);
      if (typeof body.url !== 'string' || !body.url.trim()) {
        sendJson(response, 400, { error: 'Введите ссылку на раздел со статьями' });
        return;
      }
      await pruneJobs();
      if ([...jobs.values()].some((job) => ['queued', 'running'].includes(job.status))) {
        sendJson(response, 409, { error: 'Сейчас уже выполняется другой сбор. Дождитесь его завершения' });
        return;
      }
      const limit = body.limit === 'max'
        ? Infinity
        : Math.max(1, Math.min(1000, Math.floor(Number(body.limit) || 100)));
      const dateFrom = validateDate(typeof body.dateFrom === 'string' ? body.dateFrom : '', 'с');
      let dateTo = validateDate(typeof body.dateTo === 'string' ? body.dateTo : '', 'по');
      if (dateFrom && !dateTo) dateTo = new Date().toISOString().slice(0, 10);
      if (dateFrom && dateTo && dateFrom > dateTo) {
        sendJson(response, 400, { error: 'Дата «с» должна быть раньше даты «по»' });
        return;
      }
      const job = {
        id: randomUUID(),
        status: 'queued',
        phase: 'queued',
        limit: Number.isFinite(limit) ? limit : 'max',
        pagesVisited: 0,
        pagesQueued: 0,
        catalogRead: 0,
        catalogTotal: null,
        articlesFound: 0,
        articlesProcessed: 0,
        articleFailures: 0,
        failedArticles: [],
        restrictedArticles: 0,
        restrictedArticleTitles: [],
        articlesTotal: null,
        preparedArticles: 0,
        imagesFound: 0,
        imagesIncluded: 0,
        imagesSkipped: 0,
        imagesProcessed: 0,
        imagesToProcess: 0,
        zipPercent: 0,
        undatedArticles: 0,
        crawlComplete: null,
        archivePath: null,
        archiveSize: 0,
        error: null,
        request: {
          url: body.url.trim(),
          limit: Number.isFinite(limit) ? limit : 'max',
          dateFrom,
          dateTo,
        },
        restartCount: 0,
        createdAt: Date.now(),
      };
      jobs.set(job.id, job);
      await persistJob(job);
      void runCrawlJob(job, {
        url: body.url.trim(),
        limit,
        dateFrom,
        dateTo,
      });
      sendJson(response, 202, { jobId: job.id });
    } catch (error) {
      const message = error.name === 'TimeoutError'
        ? 'Сайт отвечает слишком долго. Попробуйте позже'
        : error instanceof SyntaxError
          ? 'Некорректный запрос'
          : error.message || 'Не удалось обработать сайт';
      sendJson(response, 400, { error: message });
    }
    return;
  }

  const contentType = { '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8' }[extname(pathname)];
  if (request.method === 'GET' && pathname.startsWith('/assets/') && contentType) {
    try {
      const assetPath = join(directory, pathname.slice(1));
      const asset = await readFile(assetPath);
      response.writeHead(200, { 'Content-Type': contentType });
      response.end(asset);
    } catch {
      sendJson(response, 404, { error: 'Файл не найден' });
    }
    return;
  }

  sendJson(response, 404, { error: 'Страница не найдена' });
});

await restoreJobs();

server.listen(port, () => {
  console.log(`Site to Word запущен: http://localhost:${port}`);
});
