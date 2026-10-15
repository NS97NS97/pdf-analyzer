// ===== Элементы страницы =====
const btnPath    = document.getElementById('btnPath');
const pathBlock  = document.getElementById('pathBlock');
const folderPath = document.getElementById('folderPath');
const btnAnalyze = document.getElementById('btnAnalyze');
const btnCancel  = document.getElementById('btnCancel');
const statusBox  = document.getElementById('status');

const btnExcel       = document.getElementById('btnExcel');
const resultEmpty    = document.getElementById('resultEmpty');
const resultTableWrap= document.getElementById('resultTableWrap');
const resultBody     = document.getElementById('resultBody');

// ===== Показать / скрыть сообщение =====
function showStatus(text, type) {
    statusBox.textContent = text;
    statusBox.className = 'status ' + type;
}

function hideStatus() {
    statusBox.className = 'status hidden';
}

// ===== Проверяем, где запущен сайт: на своём компьютере или на хостинге =====
(async function checkEnvironment() {
    try {
        const data = await (await fetch('/env')).json();

        folderPath.placeholder = 'Например: ' + data.example;

        if (!data.desktop) {
            // На хостинге нет доступа к файлам компьютера — предупреждаем
            document.getElementById('hostWarning').classList.remove('hidden');
        }
    } catch (err) {
        // Не критично: страница работает и без этой проверки
    }
})();

// ===== Кнопка «Задать путь к архиву» =====
btnPath.addEventListener('click', async () => {
    btnPath.disabled = true;
    showStatus('Открываю окно выбора папки…', 'info');
    try {
        const response = await fetch('/choose_folder');
        const data = await response.json();

        if (data.fallback) {
            // Окно выбора недоступно (хостинг) — вводим путь вручную
            showStatus(data.error, 'info');
            pathBlock.classList.remove('hidden');
            folderPath.focus();
            return;
        }
        if (!response.ok) {
            showStatus(data.error || 'Не удалось открыть выбор папки.', 'error');
            pathBlock.classList.remove('hidden');
            folderPath.focus();
            return;
        }

        if (data.path) {
            folderPath.value = data.path;
            hideStatus();
            startAnalyze();              // папка выбрана — сразу анализируем
        } else {
            hideStatus();                // пользователь закрыл окно, ничего не выбрав
        }
    } catch (err) {
        // Сервер недоступен — оставляем возможность ввести путь вручную
        pathBlock.classList.remove('hidden');
        folderPath.focus();
    } finally {
        btnPath.disabled = false;
    }
});

btnCancel.addEventListener('click', () => {
    pathBlock.classList.add('hidden');
    hideStatus();
});

// ===== Запуск анализа =====
async function startAnalyze() {
    const path = folderPath.value.trim();
    if (!path) {
        showStatus('Путь к папке не указан.', 'error');
        return;
    }

    // Блокируем кнопки на время работы
    btnAnalyze.disabled = true;
    btnAnalyze.innerHTML = '<span class="spinner"></span>Анализируем…';
    showStatus('Сканирую папку и читаю PDF-файлы, подождите…', 'info');

    try {
        const response = await fetch('/analyze', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({path: path})
        });

        const data = await response.json();

        if (!response.ok) {
            showStatus(data.error || 'Ошибка при анализе.', 'error');
            return;
        }

        renderResults(data.items);

        let message = `Готово! Найдено документов: ${data.count}. Отчёт сохранён (${data.report}).`;
        if (data.report_locked) {
            message += ' Файл report.xlsx открыт в Excel, поэтому сохранена копия с датой.';
        }
        if (data.no_title > 0) {
            message += ` Не удалось распознать название у ${data.no_title} файлов (скан без текста).`;
        }
        showStatus(message, 'ok');
        btnExcel.classList.remove('disabled');   // разблокируем «Открыть xls»
        pathBlock.classList.add('hidden');

    } catch (err) {
        showStatus('Не удалось связаться с сервером: ' + err, 'error');
    } finally {
        btnAnalyze.disabled = false;
        btnAnalyze.innerHTML = '🔍 Проанализировать';
    }
}

// ===== Кнопка «Проанализировать» =====
btnAnalyze.addEventListener('click', startAnalyze);

// ===== Отрисовка таблицы результатов =====
function renderResults(items) {
    resultBody.innerHTML = '';

    if (!items.length) {
        resultEmpty.textContent = 'В указанной папке PDF-файлы не найдены.';
        resultEmpty.classList.remove('hidden');
        resultTableWrap.classList.add('hidden');
        btnExcel.classList.add('disabled');
        return;
    }

    items.forEach((item, index) => {
        const tr = document.createElement('tr');

        const num = document.createElement('td');
        num.textContent = index + 1;

        const title = document.createElement('td');
        title.textContent = item.title;

        const author = document.createElement('td');
        author.textContent = item.author;

        const year = document.createElement('td');
        year.className = 'year-cell';
        year.textContent = item.year;

        const path = document.createElement('td');
        path.className = 'path-cell';
        const link = document.createElement('a');
        link.href = '#';
        link.className = 'path-link';
        link.textContent = item.path;
        link.title = 'Открыть файл';
        link.addEventListener('click', (event) => {
            event.preventDefault();
            openFile(item.path);
        });
        path.appendChild(link);

        tr.append(num, title, author, year, path);
        resultBody.appendChild(tr);
    });

    resultEmpty.classList.add('hidden');
    resultTableWrap.classList.remove('hidden');
}

// ===== Открытие файла статьи по клику на путь =====
async function openFile(path) {
    try {
        const response = await fetch('/open_file?path=' + encodeURIComponent(path));
        const data = await response.json();

        if (!response.ok) {
            showStatus(data.error || 'Не удалось открыть файл.', 'error');
            return;
        }
        if (data.opened) {
            showStatus('Файл открыт: ' + data.file, 'ok');
        } else if (data.download) {
            // Сервер без рабочего стола (хостинг) — открыть файл на вашем
            // компьютере нельзя, поэтому он скачивается.
            showStatus(data.note || 'Файл не удалось открыть — скачиваю.', 'info');
            window.location.href = data.download;
        } else {
            showStatus('Не удалось открыть файл на этой системе: ' + data.file, 'info');
        }
    } catch (err) {
        showStatus('Ошибка при открытии файла: ' + err, 'error');
    }
}

// ===== Кнопка «Открыть xls» =====
btnExcel.addEventListener('click', async (event) => {
    event.preventDefault();

    try {
        const response = await fetch('/open_xls');
        const data = await response.json();

        if (!response.ok) {
            showStatus(data.error || 'Файл ещё не готов.', 'error');
            return;
        }

        if (data.opened) {
            showStatus('Файл report.xlsx открыт в Excel.', 'ok');
        } else if (data.download) {
            // На хостинге Excel открыть нельзя — файл скачивается
            showStatus(data.note || 'Файл отчёта скачивается.', 'info');
            window.location.href = data.download;
        }
    } catch (err) {
        window.location.href = '/report.xlsx';
    }
});
