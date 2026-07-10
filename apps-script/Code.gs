/**
 * ═══════════════════════════════════════════════════════════════
 *  ПОЛУРЫВОК 2026 — приём заявок с сайта регистрации
 *  Скрипт привязывается к Google Таблице (журналу учёта участников)
 *  Инструкция по развёртыванию — в README.md репозитория
 * ═══════════════════════════════════════════════════════════════
 */

const SHEET_NAME = 'Участники';
const SEND_CONFIRMATION = false; // письмо-подтверждение участнику (выкл; позже включим на почту организатора)
const ORGANIZER_EMAIL = '';     // укажите email — будут приходить уведомления о новых заявках (пусто = выкл.)

/* КЛЮЧ АДМИНИСТРАТОРА для скачивания полной базы участников.
   Задайте свой длинный случайный ключ ПРЯМО В РЕДАКТОРЕ Apps Script.
   ВАЖНО: никогда не публикуйте реальный ключ в GitHub — в репозитории
   должна остаться только эта заглушка. */
const ADMIN_KEY = 'СМЕНИТЕ_НА_ДЛИННЫЙ_СЛУЧАЙНЫЙ_КЛЮЧ';

const HEADERS = [
  'Дата заявки', 'ФИО', 'Дата рождения', 'Возраст на турнир', 'Пол',
  'Группа', 'Собственный вес, кг', 'Весовая категория', 'Вес гирь, кг',
  'Турнир(ы)', 'Взнос, ₽', 'Взнос оплачен',
  'Страна, город', 'Клуб', 'Email', 'Телефон', 'Почтовый адрес',
  'Согласие ПДн', 'Ответственность за здоровье', 'Публикация в списке'
];
const COL_EMAIL = 15;          // № столбца Email (для защиты от дублей)
const COL_PUBLIC = 20;         // № столбца «Публикация в списке»

/* ── Приём заявки ─────────────────────────────────────────── */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);

    // Ловушка для ботов
    if (data.website) return json_({ ok: true });

    // Обязательные поля
    const required = ['tournaments','fio','birthDate','gender','group','bodyWeight','kbWeight','city','email'];
    for (var i = 0; i < required.length; i++) {
      if (!data[required[i]]) return json_({ ok:false, error:'Не заполнено обязательное поле.' });
    }
    if (!data.consentPd || !data.consentHealth)
      return json_({ ok:false, error:'Не подтверждены обязательные согласия.' });

    const sh = getSheet_();

    // Защита от повторной заявки: тот же email
    const emails = sh.getLastRow() > 1
      ? sh.getRange(2, COL_EMAIL, sh.getLastRow() - 1, 1).getValues().flat().map(String)
      : [];
    if (emails.indexOf(String(data.email).trim()) !== -1)
      return json_({ ok:false, error:'Заявка с этой электронной почтой уже зарегистрирована. Если нужно изменить данные — напишите организатору.' });

    sh.appendRow([
      new Date(),
      String(data.fio).trim(),
      String(data.birthDate),
      data.age != null ? Number(data.age) : '',
      String(data.gender),
      Number(data.group),
      Number(data.bodyWeight),
      String(data.category || ''),
      Number(data.kbWeight),
      String(data.tournaments),
      Number(data.fee || 0),
      'нет',                              // отметку об оплате ставит организатор вручную
      String(data.city || '').trim(),
      String(data.club || '').trim(),
      String(data.email).trim(),
      String(data.phone || '').trim(),
      String(data.postAddress || '').trim(),
      data.consentPd ? 'да' : 'нет',
      data.consentHealth ? 'да' : 'нет',
      data.consentPublic ? 'да' : 'нет'
    ]);

    if (SEND_CONFIRMATION) sendConfirmation_(data);
    if (ORGANIZER_EMAIL) {
      MailApp.sendEmail(ORGANIZER_EMAIL,
        'Новая заявка: ' + data.fio,
        data.tournaments + '\nГруппа ' + data.group + ', гири ' + data.kbWeight + ' кг, ' + data.city +
        '\nВзнос: ' + data.fee + ' ₽\nEmail: ' + data.email);
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok:false, error:'Ошибка сервера: ' + String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ── Публичный список участников + выгрузка для админа ───── */
function doGet(e) {
  try {
    // Режим администратора: ?key=ВАШ_КЛЮЧ — скачивание полной базы в CSV
    if (e && e.parameter && e.parameter.key) {
      if (ADMIN_KEY !== 'СМЕНИТЕ_НА_ДЛИННЫЙ_СЛУЧАЙНЫЙ_КЛЮЧ' && e.parameter.key === ADMIN_KEY) {
        return exportCsv_();
      }
      return json_({ ok:false, error:'Неверный ключ доступа.' });
    }

    const sh = getSheet_();
    const last = sh.getLastRow();
    var participants = [];
    var total = 0;
    if (last > 1) {
      const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      total = rows.length;
      participants = rows
        .filter(function(r){ return String(r[COL_PUBLIC - 1]).toLowerCase() === 'да'; }) // согласие на публикацию
        .map(function(r){
          return {
            fio: r[1],
            city: r[12],
            group: r[5],
            kbWeight: r[8],
            category: r[7],
            gender: r[4],
            tournaments: r[9]
          };
        });
    }
    return json_({ ok:true, total: total, participants: participants });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

/* ── Служебные ────────────────────────────────────────────── */
function exportCsv_() {
  const sh = getSheet_();
  const rows = sh.getDataRange().getDisplayValues();
  // BOM + разделитель «;» — чтобы русский Excel открыл файл без танцев
  const csv = '\uFEFF' + rows.map(function(r){
    return r.map(function(c){ return '"' + String(c).replace(/"/g, '""') + '"'; }).join(';');
  }).join('\r\n');
  return ContentService
    .createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV)
    .downloadAsFile('polusnatch_2026_uchastniki.csv');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1b1a18').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sendConfirmation_(d) {
  try {
    const cat = d.category ? ('\nВесовая категория: ' + d.category) : '';
    MailApp.sendEmail({
      to: String(d.email).trim(),
      subject: 'Заявка принята — «Полурывок 2026»',
      body:
        'Здравствуйте, ' + d.fio + '!\n\n' +
        'Ваша заявка принята.\n\n' +
        'Турнир(ы): ' + d.tournaments + '\n' +
        'Группа: ' + d.group + '\n' +
        'Вес гирь: ' + d.kbWeight + ' кг' + cat + '\n\n' +
        'СТАРТОВЫЙ ВЗНОС: ' + d.fee + ' ₽\n' +
        'Перевод на карту Сбербанка 2202 2061 5850 1465 (Сергей Леонидович Р.).\n' +
        'У каждого турнира свой QR-код — оплачивайте по QR выбранного турнира (QR-коды на странице регистрации).\n' +
        'Чек об оплате отправьте в Max по номеру +7 914 567-94-82 или на rsl32@mail.ru, указав своё ФИО.\n\n' +
        'Видео выполнения упражнения (представление, взвешивание участника и гирь, выполнение — вертикально, в анфас) ' +
        'отправляйте в Max или VK Рудневу С. Л. до 23 августа 2026 года.\n\n' +
        'Если вы хотите изменить данные заявки или отозвать её — просто ответьте на это письмо.\n\n' +
        'С уважением,\nОргкомитет турнира'
    });
  } catch (e) { /* письмо не критично — заявка уже сохранена */ }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
