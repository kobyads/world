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

  if (status.phase === 'scanning') {
    const remaining = status.limit === 'max'
      ? `В очереди осталось проверить страниц: ${Math.max(0, status.pagesQueued - status.pagesVisited)}; остаток статей станет известен после обхода.`
      : `До выбранного лимита осталось: ${Math.max(0, status.limit - status.articlesFound)} статей.`;
    setProgress({
      title: 'Поиск статей',
      number: `Найдено: ${status.articlesFound} / ${status.limit === 'max' ? '…' : status.limit}`,
      detail: `Проверено страниц: ${status.pagesVisited}. ${remaining}`,
    });
    return;
  }

  if (status.phase === 'preparing-documents') {
    const total = status.articlesTotal || 0;
    const prepared = status.preparedArticles || 0;
    setProgress({
      title: 'Подготовка Word-файлов',
      number: `${prepared} / ${total}`,
      detail: `Создано документов: ${prepared} из ${total}. Обработано изображений: ${status.imagesProcessed} из ${status.imagesToProcess}.`,
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
    const link = document.createElement('a');
    link.href = `/api/jobs/${result.jobId}/download`;
    link.download = '';
    link.textContent = 'Скачать ZIP';

    message.className = 'message success';
    const notes = [];
    if (job.undatedArticles) notes.push(`Пропущено статей без распознанной даты: ${job.undatedArticles}.`);
    if (job.restrictedArticles) notes.push(`Недоступно без платного или отдельного доступа: ${job.restrictedArticles}.`);
    if (job.articleFailures) notes.push(`Не удалось выгрузить статей: ${job.articleFailures}.`);
    if (job.imagesSkipped) notes.push(`Пропущено изображений: ${job.imagesSkipped}. Подробности в отчёте внутри ZIP.`);
    if (job.crawlComplete === false && !job.articleFailures) notes.push('Обход остановился до завершения. Архив может быть неполным.');
    message.textContent = `Готово: ${job.articlesTotal} DOCX-файлов; встроено изображений: ${job.imagesFound}. ${notes.join(' ')} `;
    message.append(link);
    link.click();
    setProgress({
      title: 'Архив готов',
      number: '100%',
      detail: 'Браузер скачивает архив. Распакуйте ZIP, затем откройте DOCX в Word.',
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
