const form = document.querySelector('#export-form');
const input = document.querySelector('#url');
const limitInput = document.querySelector('#limit');
const dateFromInput = document.querySelector('#date-from');
const dateToInput = document.querySelector('#date-to');
const button = document.querySelector('#submit-button');
const message = document.querySelector('#message');
const progressPanel = document.querySelector('#progress-panel');
const progressTitle = document.querySelector('#progress-title');
const progressNumber = document.querySelector('#progress-number');
const progressMeter = document.querySelector('#progress-meter');
const progressDetail = document.querySelector('#progress-detail');

function setProgress({ title, number = '', detail = '', percent = null }) {
  progressPanel.hidden = false;
  progressTitle.textContent = title;
  progressNumber.textContent = number;
  progressDetail.textContent = detail;
  if (percent === null) {
    progressMeter.removeAttribute('value');
  } else {
    progressMeter.value = Math.max(0, Math.min(100, percent));
  }
}

function wait(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function formatBytes(value) {
  if (value < 1024 * 1024) return `${Math.max(0, Math.ceil(value / 1024))} КБ`;
  return `${(value / (1024 * 1024)).toFixed(1)} МБ`;
}

async function readJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`, { cache: 'no-store' });
  const status = await response.json();
  if (!response.ok) throw new Error(status.error || 'Не удалось получить состояние задачи');
  return status;
}

function renderJob(status) {
  if (status.phase === 'checking-rules' || status.phase === 'queued') {
    setProgress({
      title: 'Подключение к сайту',
      detail: 'Проверяем доступность сайта и правила robots.txt…',
    });
    return;
  }

  if (status.phase === 'loading-catalog') {
    const total = status.catalogTotal || 0;
    const read = status.catalogRead || 0;
    setProgress({
      title: 'Загрузка списка статей',
      number: total ? `${read} / ${total}` : '…',
      detail: total
        ? `Прочитано записей: ${read} из ${total}; доступно к выгрузке: ${status.articlesFound}; закрыто доступом: ${status.restrictedArticles || 0}. Осталось записей: ${Math.max(0, total - read)}.`
        : 'Получаем общее количество статей с сайта…',
      percent: total ? (read / total) * 100 : null,
    });
    return;
  }

  if (status.phase === 'loading-articles') {
    const total = status.articlesTotal || 0;
    const loaded = status.articlesProcessed || 0;
    setProgress({
      title: 'Загрузка текстов статей',
      number: `${loaded} / ${total}`,
      detail: `Прочитано статей: ${loaded} из ${total}; осталось: ${Math.max(0, total - loaded)}; ошибок: ${status.articleFailures || 0}.`,
      percent: total ? (loaded / total) * 100 : 0,
    });
    return;
  }

  if (status.phase === 'scanning' || status.phase === 'downloading-images') {
    const downloadingImages = status.phase === 'downloading-images';
    if (downloadingImages && status.catalogTotal !== null) {
      const total = status.articlesTotal || 0;
      const loaded = status.articlesProcessed || 0;
      setProgress({
        title: 'Загрузка изображений',
        number: `${loaded} / ${total} статей`,
        detail: `Картинок обработано: ${status.imagesProcessed} / ${status.imagesToProcess}; статей ещё предстоит прочитать: ${Math.max(0, total - loaded)}.`,
        percent: total ? (loaded / total) * 100 : 0,
      });
      return;
    }
    const remaining = status.limit === 'max'
      ? `В очереди осталось проверить страниц: ${Math.max(0, status.pagesQueued - status.pagesVisited)}; остаток статей станет известен после обхода.`
      : `До выбранного лимита осталось: ${Math.max(0, status.limit - status.articlesFound)} статей.`;
    setProgress({
      title: downloadingImages ? 'Загрузка изображений' : 'Поиск статей',
      number: `Найдено: ${status.articlesFound} / ${status.limit === 'max' ? '…' : status.limit}`,
      detail: downloadingImages
        ? `Проверено страниц: ${status.pagesVisited}; обработано изображений: ${status.imagesProcessed} / ${status.imagesToProcess}. ${remaining}`
        : `Проверено страниц: ${status.pagesVisited}. ${remaining}`,
    });
    return;
  }

  if (status.phase === 'preparing-documents') {
    const total = status.articlesTotal || 0;
    const prepared = status.preparedArticles || 0;
    setProgress({
      title: 'Подготовка Word-файлов',
      number: `${prepared} / ${total}`,
      detail: `Создано документов: ${prepared} из ${total}; осталось: ${Math.max(0, total - prepared)}; найдено изображений: ${status.imagesFound}.`,
      percent: total ? (prepared / total) * 100 : 0,
    });
    return;
  }

  if (status.phase === 'packing-zip') {
    setProgress({
      title: 'Упаковка ZIP-архива',
      number: `${status.zipPercent}%`,
      detail: `${status.articlesTotal} Word-файлов готовы; осталось упаковать: ${Math.max(0, 100 - status.zipPercent)}%.`,
      percent: status.zipPercent,
    });
    return;
  }

  if (status.phase === 'complete') {
    setProgress({
      title: 'Готово',
      number: '100%',
      detail: `Подготовлено статей: ${status.articlesTotal}; изображений: ${status.imagesFound}.`,
      percent: 100,
    });
  }
}

async function downloadArchive(response, articleCount) {
  const total = Number(response.headers.get('content-length') || 0);
  const reader = response.body?.getReader();
  if (!reader) return response.blob();

  const chunks = [];
  let received = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    received += value.byteLength;
    const remaining = total ? Math.max(0, total - received) : null;
    setProgress({
      title: 'Скачивание архива',
      number: total ? `${Math.floor((received / total) * 100)}%` : formatBytes(received),
      detail: total
        ? `Осталось скачать: ${formatBytes(remaining)} из ${formatBytes(total)}; статей: ${articleCount}.`
        : `Скачано: ${formatBytes(received)}; размер архива пока неизвестен.`,
      percent: total ? (received / total) * 100 : null,
    });
  }

  return new Blob(chunks, {
    type: response.headers.get('content-type') || 'application/zip',
  });
}

async function waitForCompletion(jobId) {
  while (true) {
    const status = await readJob(jobId);
    renderJob(status);
    if (status.status === 'error') throw new Error(status.error || 'Не удалось собрать статьи');
    if (status.status === 'done') return status;
    await wait(800);
  }
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  message.className = 'message';
  message.textContent = 'Запускаем сбор…';
  button.disabled = true;
  setProgress({ title: 'Запуск задачи', detail: 'Подключаемся к сайту…' });

  try {
    const response = await fetch('/api/crawl', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url: input.value.trim(),
        limit: limitInput.value,
        dateFrom: dateFromInput.value,
        dateTo: dateToInput.value,
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Не удалось запустить сбор');

    const job = await waitForCompletion(result.jobId);
    setProgress({
      title: 'Скачивание архива',
      number: `${job.articlesTotal} статей`,
      detail: 'Архив подготовлен, начинаем скачивание…',
    });
    const downloadResponse = await fetch(`/api/jobs/${result.jobId}/download`);
    if (!downloadResponse.ok) {
      const error = await downloadResponse.json();
      throw new Error(error.error || 'Не удалось скачать архив');
    }

    const blob = await downloadArchive(downloadResponse, job.articlesTotal);
    const disposition = downloadResponse.headers.get('content-disposition') || '';
    const filename = disposition.match(/filename="([^"]+)"/i)?.[1] || 'articles.zip';
    const articleCount = downloadResponse.headers.get('x-articles-count') || job.articlesTotal;
    const imageCount = Number(downloadResponse.headers.get('x-images-count') || 0);
    const articleFailures = Number(downloadResponse.headers.get('x-article-failures') || 0);
    const restrictedCount = Number(downloadResponse.headers.get('x-restricted-articles') || 0);
    const undatedCount = Number(downloadResponse.headers.get('x-undated-articles') || 0);
    const complete = downloadResponse.headers.get('x-crawl-complete') !== 'false';
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 60_000);

    message.className = 'message success';
    const notes = [];
    if (undatedCount) notes.push(`Пропущено статей без распознанной даты: ${undatedCount}.`);
    if (restrictedCount) notes.push(`Недоступно без платного/соответствующего доступа: ${restrictedCount}.`);
    if (articleFailures) notes.push(`Не удалось прочитать статей: ${articleFailures}.`);
    if (!complete && !articleFailures) notes.push('Обход остановился до завершения — архив может быть неполным.');
    message.textContent = `Готово: ${articleCount} DOCX-файлов; встроено изображений: ${imageCount}.${notes.length ? ` ${notes.join(' ')}` : ''}`;
    setProgress({
      title: 'Архив скачан',
      number: '100%',
      detail: 'Распакуйте ZIP, затем открывайте DOCX-файлы в Word.',
      percent: 100,
    });
  } catch (error) {
    message.className = 'message error';
    message.textContent = error.message || 'Не удалось обработать страницу';
    setProgress({ title: 'Сбор остановлен', detail: error.message || 'Неизвестная ошибка' });
  } finally {
    button.disabled = false;
  }
});
