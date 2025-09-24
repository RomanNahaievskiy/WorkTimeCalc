// запуск інтерфейсу
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("ui") // ім'я твого HTML-файлу в проекті
    .setTitle("Облік робочого часу")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// Визначення ід цільових таблиць
let dbId = '1-JG_W71qHXcuq6h8Ur7nZpNZVRJBlf9XqHMxNoM1wuM'
let wlId = '1-EAVYT9PiEq-BKL_BrikRANynPewdo6mDxKBtKvgx1o'

function getEmployees() {
  return SpreadsheetApp.openById(dbId).getSheetByName("db1").getDataRange().getValues();
}

function getJournal() {
  return SpreadsheetApp.openById(wlId).getSheetByName("Журнал обліку та відвідування");
}

// ======================== Допоміжні функції (час) ======================= //
function getTimeHhMmSsStr(ts) {
  // Отримати час у форматі hh:mm:ss
  const hours = ts.getHours().toString().padStart(2, "0"); // Години
  const minutes = ts.getMinutes().toString().padStart(2, "0"); // Хвилини
  const seconds = ts.getSeconds().toString().padStart(2, "0"); // Секунди
  return (time = `${hours}:${minutes}:${seconds}`);
}
function getDayDdMmYyyyStr(ts) {
  // Отримати дату у форматі dd.mm.yyyy
  const day = ts.getDate().toString().padStart(2, "0"); // День
  const month = (ts.getMonth() + 1).toString().padStart(2, "0"); // Місяць (0-indexed)
  const year = ts.getFullYear(); // Рік

  return (date = `${day}.${month}.${year}`);
}

// Функція для перетворення часу у форматі HH:MM:SS у секунди
function timeToSeconds(timeString) {
  const [hours, minutes, seconds] = timeString.split(":").map(Number);
  return hours * 3600 + minutes * 60 + seconds;
}

// Функція для перетворення секунд у формат HH:MM:SS
function secondsToTime(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remainingSeconds = seconds % 60;

  return [
    hours.toString().padStart(2, "0"),
    minutes.toString().padStart(2, "0"),
    remainingSeconds.toString().padStart(2, "0"),
  ].join(":");
}

// Визначення пропрацьованого часу
function worktime(employeId, shiftType) {
  const timestamp = new Date();
  // якщо shiftType == "Кінець зміни"
  // логіка для визначення відпрацьованого часу worktime
  if (shiftType == "Кінець зміни") {
    // потрібно знайти в журналі останній запис про початок зміни
    const lastLogStartShift = onShift(employeId, "Початок зміни");
    //  отримати з нього мітку часу

    // Обчислюємо різницю у мілісекундах
    let diff = timestamp - lastLogStartShift.timeStamp;

    //
    // Перетворюємо мілісекунди в години, хвилини, секунди
    let hours = Math.floor(diff / (1000 * 60 * 60)); // Кількість годин
    let minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60)); // Залишок хвилин
    let seconds = Math.floor((diff % (1000 * 60)) / 1000); // Залишок секунд

    // Форматуємо в [hh:mm:ss]
    let worktimeTotal = `${hours.toString().padStart(2, "0")}:${minutes
      .toString()
      .padStart(2, "0")}:${seconds.toString().padStart(2, "0")}`;

    return worktimeTotal;
  }
}
// тривалість перерви
function getBreakDuration(worktimeTotal) {
  const worktimeInSeconds = Number(timeToSeconds(worktimeTotal));

  let breakDuration; // Змінна для зберігання тривалості перерви

  // Використання switch з true
  switch (true) {
    case worktimeInSeconds < 2 * 60 * 60: // Менше 15 хв
      breakDuration = "00:00:00"; // 0 хвилин
      break;

    case worktimeInSeconds > 2 * 60 * 60 + 15 * 60 &&
      worktimeInSeconds < 4 * 60 * 60: // Більше 2: 15 і Менше 4 годин
      breakDuration = "00:15:00"; // 15 хвилин
      break;

    case worktimeInSeconds >= 4 * 60 * 60 && worktimeInSeconds < 7 * 60 * 60: // Від 4 до 7 годин
      breakDuration = "00:30:00"; // 30 хвилин
      break;

    case worktimeInSeconds >= 7 * 60 * 60 && worktimeInSeconds <= 12 * 60 * 60: // Від 7 до 12 годин
      breakDuration = "01:00:00"; // 1 година
      break;

    default: // Якщо час більше 12 годин або менше 0
      breakDuration = "Зміна > 12 год!"; // Дефолтний випадок
  }

  return breakDuration;
}
// Вирахувати робочий час
function getWorkingHours(worktimeTotal, breakDur) {
  if (breakDur !== "Зміна > 12 год!") {
    const worktimeInSeconds = timeToSeconds(worktimeTotal);
    const breakInSeconds = timeToSeconds(breakDur);

    // Робочий час = загальний час - перерва

    const workingTimeInSeconds = worktimeInSeconds - breakInSeconds;

    // Перевірка на коректність часу
    if (workingTimeInSeconds < 0) {
      return "Невірний час! Робочий час не може бути меншим за тривалість перерви. Хтось редагував журннал!";
    }
    const workingHours = secondsToTime(workingTimeInSeconds);

    return workingHours;
  }
  return "???";
}

// Спрацює лише при незакритій вчасно зміні  більше 12 год
function getNotice(employeId, shiftType, employeeName) {
  const lastJournalentry = onShift(employeId, shiftType);

  const notice = `Перевір дані реєстрації працівника ${employeeName}, що не закрив зміну, яка  розпочалася ${getDayDdMmYyyyStr(
    lastJournalentry.timeStamp
  )} о ${getTimeHhMmSsStr(lastJournalentry.timeStamp)} ( № запису ${
    lastJournalentry.row + 1
  })!`;
  // + 1 щоб виправити зміщення діапазону  в журналі ??
  return notice;
}

//ISO8601 вирахувати номер тижня в році
function getISOWeekNum(date) {
  // date = new Date(2024, 11, 27)
  const tempDate = new Date(date.getFullYear(), 0, 4); // 4 січня поточного року
  const firstMonday = new Date(
    tempDate.setDate(tempDate.getDate() - ((tempDate.getDay() + 6) % 7))
  ); // Перший понеділок
  const diff = date - firstMonday; // Різниця у мілісекундах

  // Базовий розрахунок номера тижня
  let weekNum = Math.ceil(diff / (7 * 86400000));

  // Перевіряємо, чи поточний рік має 53 тижні
  const has53Weeks =
    new Date(date.getFullYear(), 0, 1).getDay() === 4 ||
    new Date(date.getFullYear(), 11, 31).getDay() === 4;

  // Додаємо +1 лише якщо рік має 53 тижні і обчислений тиждень більше 52
  if (has53Weeks && weekNum > 52) {
    weekNum += 1;
  }

  return weekNum;
}
// Отримати номер дня в році deprecated
// function getISODayNum(date) {
//   const startOfYear = new Date(date.getFullYear(), 0, 1); // Створюємо дату для початку поточного року
//   const dayOfYear = Math.ceil((date - startOfYear + 1) / (1000 * 60 * 60 * 24)); // Розраховуємо день року
//   return dayOfYear;
// }
// Отримати назву дня тижня
function getDayName(date) {
  // Створити об'єкт Date (якщо date — це рядок або інше представлення дати)
  const inputDate = new Date(date);

  // Масив назв днів тижня
  const days = ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];

  // Отримати день тижня (0 — неділя, 1 — понеділок, і т.д.)
  const dayIndex = inputDate.getDay();

  // Повернути назву дня
  return days[dayIndex];
}

// ============================= ОТРИМАННЯ ІНФОРМАЦІЇ ІЗ БД - ІНШОЇ ТАБЛИЦІ=========================== //

// Парсинг ID таблиці (url- таблиці)
function getSpreadsheetIdFromUrl(url) {
  var regex = /\/d\/([a-zA-Z0-9_-]+)\//;
  var match = url.match(regex);

  if (match) {
    return match[1]; // Повертає ID, якщо воно знайдене
  } else {
    throw new Error("ID не знайдено в URL");
  }
}
// Отримати таблицю по url
function getTargetSheet(url) {
  var id = getSpreadsheetIdFromUrl(url); // з глобальної обл . допоміжних функцій парсить id

  const targetSheet = SpreadsheetApp.openById(id);
  return targetSheet;
}
// Отримати адресу таб із (Аркуша, клітинки)
function getUrlSheet(nameSheet = "Налаштування БД", cell = "A6") {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  if (!sheet) {
    throw new Error("Лист з назвою " + nameSheet + " не знайдено");
  }

  var urlSDB = sheet.getRange(cell).getValue();
  if (!urlSDB) {
    throw new Error("URL Таблиці не знайдено в клітинці " + cell);
  }
  return urlSDB;
}
// функція читання даних із таблиці <readonly> повертає двовимірний масив [[row0],[row1],...]
function getSheetData(
  targetSheet,
  nameSheetDB = "Журнал обліку та відвідування"
) {
  try {
    if (!targetSheet) {
      throw new Error("Таблицю " + targetSheet + " не знайдено");
    }
    if (!nameSheetDB) {
      throw new Error("Лист " + nameSheetDB + " не знайдено");
    }
    return targetSheet.getSheetByName(nameSheetDB).getDataRange().getValues();
  } catch (error) {
    Logger.log("Err " + error.message);
  }
}

// Імпорт БД (назва листа із адресою БД, адреса клітинки із адресою БД, назва листа в сторонньому файлі БД(її потрібно знати))

function getDataDB(nameSheet, cell, nameSheetDB) {
  try {
    // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
    // if (!sheet) {
    //   throw new Error("Лист з назвою " + nameSheet + " не знайдено");
    // }

    // var urlSDB = sheet.getRange(cell).getValue();
    // if (!urlSDB) {
    //   throw new Error("URL БД не знайдено в клітинці " + cell);
    // }
    // Отримуємо ID таблиці з URL
    // var id = getSpreadsheetIdFromUrl(urlSDB);
    var id = dbId; //прямо вказую ід БД

    // Читаємо дані з іншого Google Sheets документа
    var spreadsheet = SpreadsheetApp.openById(id);
    var dataSheet = spreadsheet.getSheetByName(nameSheetDB); // Назва листа в документі
    if (!dataSheet) {
      throw new Error("Помилка в назві листа БД  " + nameSheetDB);
    }
    var data = dataSheet.getDataRange().getValues();
    // повертаємо дані

    return data;
  } catch (error) {
    // Логування помилки
    Logger.log("Помилка: " + error.message);
    return null; // Можна повернути null або порожній масив, якщо сталася помилка
  }
}
// ===================== Порівняння даних ============================//
//  додати цей виклик в google.scripts.run
function compareData(employeId, shiftType) {
  let employeeName = "";
  let isValid = false;
  let greting =
    shiftType === "Початок зміни"
      ? "Продуктивної праці "
      : "Гарного відпочинку "; // якщо початок то Продуктивної праці інакше Гарного відпочинку
  try {
    // Отримуємо дані з іншої таблиці (потребує дозволів ) test!!!
    const dbData = getEmployees().filter((row) =>
      row.some((cell) => cell !== "")
    );

    // Порівнюємо дані із базою даних
    isValid = dbData.some((row) => row[0] === employeId);
    if (isValid) {
      employeeName = dbData.find((row) => row[0] === employeId)[1]; // [1] - це тому , що ID в колонці [0], а ім'я знаходиться в колонці з індексом 1
    }

    // тут додаткова валідація

    // Якщо працівника не знайдено в БД, повертаємо помилку
    if (!isValid) {
      return {
        isValid: false,
        message: "Дані не знайдено в БД.",
        name: employeeName,
      };
    }

    // Перевіряємо запис у журналі (останній запис цього працівника)
    const journalEntry = onShift(employeId, shiftType);

    if (journalEntry == undefined && isValid) {
      // Якщо працівник вперше реєструється в журналі');
      if (shiftType === "Кінець зміни") {
        // Повертаємо на фронтенд успішний результат, якщо всі перевірки пройдені
        return {
          isValid: false,
          message: `Працівник ${employeeName}, Вітаю, ви тут вперше! Оберіть "Початок зміни"! `,
          name: employeeName,
        };
      }

      // Повертаємо успішний результат, якщо всі перевірки пройдені
      return {
        isValid: true,
        message: `${employeeName}, Вітаю! ${greting}.`,
        name: employeeName,
      };
    } else if (journalEntry.typeEntry === shiftType) {
      return {
        isValid: false,
        message: `${employeeName} вже здійснив реєстрацію типу ${shiftType}!`,
      };
    }

    // Повертаємо успішний результат, якщо всі перевірки пройдені
    return {
      isValid: true,
      message: `${greting}, ${employeeName}`,
      name: employeeName,
    };
  } catch (error) {
    Logger.log("Помилка: " + error.message);
    return {
      isValid: false,
      message: "Невалідний ID " + error.message,
    };
  }
}
// =========================== VALIDATION ONSHIFT ======================================//

// Отримуємо дані з журналу таблиці
function onShift(employeId = 'id123', shiftType) {
  let id = employeId;
  let journalEntry;
  // отримуємо доступ до журналу -- Налаштувати!!!
  const logger = getJournal().getDataRange().getValues().filter((row) => row.some((cell) => cell !== ""));

  // Пошук в журналі з кінця до першого збігу
  function findJournalEntry(logger, id) {
    if (id) {
      for (let i = logger.length - 1; i >= 0; i--) {
        // Шукаємо значення в поточному рядку
        if (logger[i].includes(id)) {
          return {
            row: i,
            column: logger[i].indexOf(id),
            value: id,
            typeEntry: logger[i][1],
            timeStamp: logger[i][4], // мітка часу для обчислення Якщо в журналі зміниться колонка - змінити й тут
          };
        }
      }

      // Якщо значення не знайдено
      return undefined;
    }
  }

  journalEntry = findJournalEntry(logger, id);
  Logger.log(journalEntry)
  return journalEntry;
}
// =========================== ФОРМА АВТОРИЗАЦІЇ ПРАЦІВНИКА  =========================== //
// Функції для інтерфейсу Обліку обігу терміналів ДЛЯ ГУГЛ ТАБЛИЦІ

// // показати HTML форму для обліку обігу терміналів Як шаблон до якого можна включати інші частини
// function showUiForm() {
//   const htmlForm = HtmlService.createTemplateFromFile("ui") // ⬅️ ВАЖЛИВО!
//     .evaluate()
//     .setWidth(1200)
//     .setHeight(600);

//   SpreadsheetApp.getUi().showModalDialog(htmlForm, "Облік робочого часу");
// }

// // коли документ відкрито створююменю "Авторизація" - тригер
// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu("Авторизація") // Додає нове меню - щоб почати використовувати форму авторизації
//     .addItem("Відкрити HTML", "showUiForm")
//     .addToUi();
// }
//================================ WRITE SHIFT DATA===============================================
// запис даних із форми
function writeShiftData(shiftType, employeId, employeeName) {
  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Отримуємо активний аркуш - можна налаштувати куди конкретно

  const sheet = getJournal(); //Вказати клітинку із URL файлу журналу

  const timestamp = new Date(); // Поточний час
  // Отримати час у форматі hh:mm:ss

  const time = getTimeHhMmSsStr(timestamp);
  // Отримати дату у форматі dd.mm.yyyy
  const date = getDayDdMmYyyyStr(timestamp);
  const worktimeTotal = worktime(employeId, shiftType)
    ? worktime(employeId, shiftType)
    : "";

  // якщо в працівника є загальний робочий час, то вирахувати триваліть перерви
  const breakDur =
    worktimeTotal !== "" && worktimeTotal
      ? getBreakDuration(worktimeTotal)
      : "";

  // вирахувати кількість оплачуваного робочого часу
  const worktimeShift =
    worktimeTotal !== "" && breakDur !== ""
      ? getWorkingHours(worktimeTotal, breakDur)
      : "";
  //  Примітка "Потрібно перевірити дані реєстрації праціника"
  // if (timeToSeconds(worktimeTotal) > 12 * 60 * 60) {
  //   let notice = getNotice(employeId, shiftType, employeeName);
  // }
  let notice =
    timeToSeconds(worktimeTotal) > 12 * 60 * 60
      ? "Потребує уточнення, або коригування"
      : "";

  // Номер тижня (кількість тижнів в поточному році)
  let numweek = getISOWeekNum(timestamp)
    ? getISOWeekNum(timestamp)
    : "Err numweek";

  //
  let numday = getDayName(timestamp) ? getDayName(timestamp) : "Err numday";
  //  дані відображатимуться в такому порядку стовпців:
  const formData = [
    employeId,
    shiftType,
    date,
    time,
    timestamp,
    numweek,
    numday,
    employeeName,
    worktimeTotal,
    breakDur,
    worktimeShift,
    notice,
  ]; // Масив даних для запису

  // Додаємо дані в наступний вільний рядок таблиці
  sheet.appendRow(formData);
}
