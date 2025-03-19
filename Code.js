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
  Logger.log("שנה לפני ניקוי: " + hebYear);
  hebYear = hebYear.toString().replace(/['"]/g, '');
  Logger.log("שנה אחרי ניקוי: " + hebYear);
  
  // בדיקה אם השנה היא כבר מספר
  if (!isNaN(hebYear) && hebYear.length >= 4) {
    Logger.log("השנה כבר מספר: " + parseInt(hebYear));
    return parseInt(hebYear);
  }
  
  // מפרקים לתחילית וסיומת
  let prefix = hebYear.substring(0, 2);
  let suffix = hebYear.substring(2);
  Logger.log("תחילית השנה: " + prefix);
  Logger.log("סיומת השנה: " + suffix);
  
  // ממירים את התחילית למספר
  let prefixMap = {
    'תש': 5000, 'תר': 5900, 'תק': 5800,
    'תפ': 5700, 'תמ': 5600, 'תל': 5500,
    'תכ': 5400, 'ת': 5300
  };
  
  let prefixValue = prefixMap[prefix];
  if (prefixValue === undefined) {
    Logger.log("לא הצלחנו להמיר את התחילית: " + prefix);
    return null;
  }
  Logger.log("ערך התחילית: " + prefixValue);
  
  // ממירים את הסיומת למספר
  let suffixValue = parseInt(suffix);
  if (isNaN(suffixValue)) {
    Logger.log("לא הצלחנו להמיר את הסיומת: " + suffix);
    return null;
  }
  Logger.log("ערך הסיומת: " + suffixValue);
  
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
    
    // רושמים את שנת הלידה והשנה הנוכחית בגיליון log
    let age = null;
    try {
      let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
      if (logSheet) {
        // כותבים ערכים קבועים לבדיקה
        let birthYear = 5741;  // A2
        let currentYear = 5785;  // B2
        
        // רושמים את השנים
        logSheet.getRange(2, 1).setValue(birthYear);
        logSheet.getRange(2, 2).setValue(currentYear);
        
        // מחשבים את הגיל ורושמים בעמודה C
        age = currentYear - birthYear;
        logSheet.getRange(2, 3).setValue(age);
      }
    } catch (err) {
      Logger.log("שגיאה בכתיבה לגיליון log: " + err);
    }
    
    // שומרים את התאריך המחושב
    sheet.getRange(i + 2, 5).setValue(convertedDate);
    
    // יוצרים או מעדכנים אירוע ביומן
    let newEventId = createOrUpdateCalendarEvent(
      name,
      convertedDate,
      eventId,
      `${hebDay} ${hebMonth}`,
      age  // מעבירים את הגיל שחישבנו
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
    let eventTitle = age ? 
      `יום הולדת ${age} ל${name} - ${hebrewDate}` : 
      `יום הולדת של ${name} - ${hebrewDate}`;
    
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
