/**
 * Синхронизирует события из Google Календаря в Google Таблицу
 */
function syncCalendarToSheet() {
  // Настройки (можно изменить)
  var calendarId = 'primary'; // 'primary' или email календаря
  var sheetName = 'Календарь'; // Название листа
  var daysToSync = 60; // Сколько дней вперед включать
  
  // Получаем календарь и лист
  var calendar = CalendarApp.getCalendarById(calendarId);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  
  // Диапазон дат
  var today = new Date();
  var futureDate = new Date();
  futureDate.setDate(today.getDate() + daysToSync);
  
  // Получаем события
  var events = calendar.getEvents(today, futureDate);
  
  // Подготавливаем данные
  var headers = ['ID', 'Название', 'Описание', 'Начало', 'Конец', 'Место', 'Создано', 'Статус'];
  var data = [headers];
  
  // Формируем строки
  events.forEach(function(event) {
    data.push([
      event.getId().split('@')[0], // Укорачиваем ID
      event.getTitle(),
      event.getDescription(),
      event.getStartTime(),
      event.getEndTime(),
      event.getLocation(),
      event.getDateCreated(),
      event.getEventSeries() ? 'Серия' : 'Одно'
    ]);
  });
  
  // Записываем в таблицу с проверкой существующих данных
  if (sheet.getLastRow() < 2) {
    // Если лист пустой, записываем все
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  } else {
    // Обновляем существующие и добавляем новые события
    var existingData = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat();
    
    data.slice(1).forEach(function(row) {
      var eventId = row[0];
      var rowIndex = existingData.indexOf(eventId) + 2;
      
      if (rowIndex > 1) {
        // Обновляем существующую запись
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
      } else {
        // Добавляем новую запись
        sheet.appendRow(row);
      }
    });
  }
  
  // Форматирование
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
  
  // Добавляем меню
  createMenu();
}

/**
 * Создает меню в таблице
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Календарь')
    .addItem('Синхронизировать', 'syncCalendarToSheet')
    .addToUi();
}

/**
 * При открытии таблицы
 */
function onOpen() {
  createMenu();
}