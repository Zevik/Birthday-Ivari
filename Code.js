/**
 * ממיר תאריך עברי ללועזי באמצעות Hebcal API
 */
function convertHebrewToGregorian(hebDay, hebMonth, hebYear) {
  try {
    // מנקים את הנתונים
    hebDay = hebDay.toString().replace(/['"]/g, '');
    hebMonth = hebMonth.toString().replace(/['"]/g, '');
    hebYear = hebYear ? hebYear.toString().replace(/['"]/g, '') : null;
    
    // ממירים את היום למספר
    let day = parseInt(hebDay);
    if (isNaN(day)) return null;
    
    // ממירים את החודש לפורמט Hebcal
    let monthMap = {
      'ניסן': 'Nisan', 'אייר': 'Iyar', 'סיון': 'Sivan',
      'תמוז': 'Tamuz', 'אב': 'Av', 'אלול': 'Elul',
      'תשרי': 'Tishrei', 'חשון': 'Cheshvan', 'כסלו': 'Kislev',
      'טבת': 'Tevet', 'שבט': 'Shvat', 'אדר': 'Adar'
    };
    
    let month = monthMap[hebMonth];
    if (!month) return null;
    
    // בונים URL ל-API
    let url = `https://www.hebcal.com/converter?cfg=json&hy=${hebYear || getCurrentHebrewYear()}&hm=${month}&hd=${day}&h2g=1`;
    
    // קוראים ל-API
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    
    // בודקים אם התאריך כבר עבר השנה
    let date = new Date(data.gy, data.gm - 1, data.gd);
    if (date < new Date()) {
      // ממירים לשנה הבאה
      url = `https://www.hebcal.com/converter?cfg=json&hy=${getCurrentHebrewYear() + 1}&hm=${month}&hd=${day}&h2g=1`;
      response = UrlFetchApp.fetch(url);
      data = JSON.parse(response.getContentText());
    }
    
    return new Date(data.gy, data.gm - 1, data.gd);
  } catch (err) {
    Logger.log("שגיאה בהמרת תאריך: " + err);
    return null;
  }
}

/**
 * מקבל את השנה העברית הנוכחית
 */
function getCurrentHebrewYear() {
  try {
    let today = new Date();
    let url = `https://www.hebcal.com/converter?cfg=json&gy=${today.getFullYear()}&gm=${today.getMonth() + 1}&gd=${today.getDate()}&g2h=1`;
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    return data.hy;
  } catch (err) {
    Logger.log("שגיאה בקבלת שנה עברית: " + err);
    return null;
  }
}

/**
 * ממיר שנה עברית למספר
 */
function convertHebrewYearToNumber(hebYear) {
  if (!hebYear) return null;
  
  // מנקים את השנה
  hebYear = hebYear.toString().replace(/['"]/g, '');
  
  // מפרקים לתחילית וסיומת
  let prefix = hebYear.substring(0, 2);
  let suffix = hebYear.substring(2);
  
  // ממירים את התחילית למספר
  let prefixMap = {
    'תש': 700, 'תר': 600, 'תק': 500,
    'תפ': 400, 'תמ': 300, 'תל': 200,
    'תכ': 100, 'ת': 0
  };
  
  let prefixValue = prefixMap[prefix];
  if (prefixValue === undefined) {
    Logger.log("לא הצלחנו להמיר את התחילית: " + prefix);
    return null;
  }
  
  // ממירים את הסיומת למספר
  let suffixValue = parseInt(suffix);
  if (isNaN(suffixValue)) {
    Logger.log("לא הצלחנו להמיר את הסיומת: " + suffix);
    return null;
  }
  
  let result = prefixValue + suffixValue;
  Logger.log("המרת שנה עברית למספר: " + hebYear + " -> " + result);
  return result;
}

/**
 * הפונקציה המרכזית שמעדכנת את ימי ההולדת
 */
function updateBirthdays() {
  // מקבלים את הגיליון
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ימי הולדת");
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // מקבלים את הנתונים
  let data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // עוברים על כל שורה
  for (let i = 0; i < data.length; i++) {
    let [name, hebDay, hebMonth, hebYear, gregorian, eventId] = data[i];
    
    // מדלגים על שורה ריקה
    if (!name || !hebDay || !hebMonth) continue;
    
    // ממירים את התאריך העברי ללועזי
    let convertedDate = convertHebrewToGregorian(hebDay, hebMonth, hebYear);
    if (!convertedDate) continue;
    
    // שומרים את התאריך המחושב
    sheet.getRange(i + 2, 5).setValue(convertedDate);
    
    // יוצרים או מעדכנים אירוע ביומן
    let newEventId = createOrUpdateCalendarEvent(
      name,
      convertedDate,
      eventId,
      `${hebDay} ${hebMonth}`,
      null  // מעבירים null במקום age
    );
    
    if (newEventId && newEventId !== eventId) {
      sheet.getRange(i + 2, 6).setValue(newEventId);
    }
  }
}

/**
 * יוצר או מעדכן אירוע ביומן
 */
function createOrUpdateCalendarEvent(name, dateObj, eventId, hebrewDate, age) {
  try {
    let calendar = CalendarApp.getDefaultCalendar();
    let eventTitle = `יום הולדת של ${name} - ${hebrewDate}`;  // תמיד משתמשים בכותרת ללא גיל
    
    let event = eventId ? calendar.getEventById(eventId) : null;
    
    if (event) {
      event.setTitle(eventTitle);
      event.setAllDayDate(dateObj);
      return eventId;
    } else {
      event = calendar.createAllDayEvent(eventTitle, dateObj);
      return event.getId();
    }
  } catch (err) {
    Logger.log("שגיאה ביצירה/עדכון אירוע: " + err);
    return "";
  }
}
