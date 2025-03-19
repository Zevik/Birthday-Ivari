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
    
    // מקבלים את השנה העברית הנוכחית
    let currentHebrewYear = getCurrentHebrewYear();
    
    // --- שימוש בתאריך קשיח לבדיקה ---
    // תאריך היום ב-19 במרץ, 2025
    let today = new Date(2025, 2, 19); // חודשים מתחילים מ-0 בג'אווה סקריפט
    Logger.log("[FIXED] התאריך הנוכחי: " + today.toLocaleDateString());
    
    // --- המרת התאריך העברי ללועזי בשנה העברית הנוכחית ---
    let url = `https://www.hebcal.com/converter?cfg=json&hy=${currentHebrewYear}&hm=${month}&hd=${day}&h2g=1`;
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    
    // יוצרים את התאריך מהתשובה של ה-API
    let dateThisYear = new Date(data.gy, data.gm - 1, data.gd);
    Logger.log("[FIXED] התאריך לשנה הנוכחית: " + dateThisYear.toLocaleDateString());
    
    // --- בדיקה אם התאריך כבר עבר השנה ---
    // השוואת מילישניות מאז 1.1.1970 ללא שעות
    let todayWithoutTime = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    let dateThisYearWithoutTime = new Date(dateThisYear.getFullYear(), dateThisYear.getMonth(), dateThisYear.getDate()).getTime();
    
    Logger.log("[FIXED] מילישניות של היום: " + todayWithoutTime);
    Logger.log("[FIXED] מילישניות של התאריך השנה: " + dateThisYearWithoutTime);
    Logger.log("[FIXED] האם התאריך כבר עבר? " + (dateThisYearWithoutTime < todayWithoutTime));
    
    // אם התאריך כבר עבר, נמיר לשנה הבאה
    if (dateThisYearWithoutTime < todayWithoutTime) {
      Logger.log("[FIXED] התאריך כבר עבר! ממיר לשנה הבאה...");
      
      // קוראים ל-API פעם נוספת עם השנה העברית הבאה
      url = `https://www.hebcal.com/converter?cfg=json&hy=${currentHebrewYear + 1}&hm=${month}&hd=${day}&h2g=1`;
      response = UrlFetchApp.fetch(url);
      data = JSON.parse(response.getContentText());
      
      // יוצרים את התאריך החדש מהתשובה של ה-API
      let dateNextYear = new Date(data.gy, data.gm - 1, data.gd);
      Logger.log("[FIXED] התאריך לשנה הבאה: " + dateNextYear.toLocaleDateString());
      
      return dateNextYear;
    } else {
      Logger.log("[FIXED] התאריך עוד לא עבר, משתמש בתאריך השנה הנוכחית");
      return dateThisYear;
    }
  } catch (err) {
    Logger.log("[FIXED] שגיאה בהמרת תאריך: " + err);
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
    'תש': 700, 'תר': 600, 'תק': 500,
    'תפ': 400, 'תמ': 300, 'תל': 200,
    'תכ': 100, 'ת': 0
  };
  
  let prefixValue = prefixMap[prefix];
  if (prefixValue === undefined) {
    Logger.log("לא הצלחנו להמיר את התחילית: " + prefix);
    return null;
  }
  Logger.log("ערך התחילית: " + prefixValue);
  
  // ממירים את הסיומת למספר
  let suffixValue = convertHebrewLettersToNumber(suffix);
  if (suffixValue === null) {
    Logger.log("לא הצלחנו להמיר את הסיומת: " + suffix);
    return null;
  }
  Logger.log("ערך הסיומת: " + suffixValue);
  
  // מחברים את המספרים ומוסיפים 5000 לקבלת השנה העברית המלאה
  let result = 5000 + prefixValue + suffixValue;
  Logger.log("המרת שנה עברית למספר: " + hebYear + " -> " + result);
  return result;
}

/**
 * ממיר אותיות עבריות למספר
 */
function convertHebrewLettersToNumber(letters) {
  if (!letters) return null;
  
  // מפה של אותיות עבריות לערכים מספריים
  const hebrewLetterValues = {
    'א': 1, 'ב': 2, 'ג': 3, 'ד': 4, 'ה': 5, 'ו': 6, 'ז': 7, 'ח': 8, 'ט': 9,
    'י': 10, 'כ': 20, 'ל': 30, 'מ': 40, 'נ': 50, 'ס': 60, 'ע': 70, 'פ': 80, 'צ': 90,
    'ק': 100, 'ר': 200, 'ש': 300, 'ת': 400
  };
  
  // בדיקה אם זה כבר מספר
  if (!isNaN(letters)) {
    return parseInt(letters);
  }
  
  let value = 0;
  
  // מעבר על כל אות והמרתה למספר
  for (let i = 0; i < letters.length; i++) {
    const letter = letters[i];
    const letterValue = hebrewLetterValues[letter];
    
    if (letterValue === undefined) {
      Logger.log("אות לא מוכרת: " + letter);
      return null;
    }
    
    value += letterValue;
  }
  
  return value;
}

/**
 * מחשב את הגיל על פי שנים עבריות
 */
function calculateHebrewAge(birthYear, currentYear) {
  if (!birthYear || !currentYear) return null;
  
  // חישוב ההפרש בין השנים
  let age = currentYear - birthYear;
  
  // וידוא שהגיל הגיוני (בין 0 ל-120)
  if (age < 0 || age > 120) {
    Logger.log("חישוב הגיל נתן תוצאה לא הגיונית: " + age + ". בודק אפשרויות אחרות...");
    
    // אם שנת הלידה גדולה מ-5000, ייתכן שהיא כבר כוללת את ה-5000
    if (birthYear > 5000 && currentYear > 5000) {
      age = currentYear - birthYear;
    } else if (birthYear < 1000 && currentYear > 5000) {
      // אם שנת הלידה קטנה מ-1000, אולי צריך להוסיף 5000
      age = currentYear - (birthYear + 5000);
    }
    
    // בדיקה שוב שהגיל הגיוני
    if (age < 0 || age > 120) {
      Logger.log("עדיין לא הצלחנו לחשב גיל הגיוני. מחזיר null.");
      return null;
    }
  }
  
  Logger.log("חישוב גיל סופי: " + currentYear + " - " + birthYear + " = " + age);
  return age;
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
    
    // לוג פשוט להבנת הנתונים
    Logger.log("[DEBUG] מעבד שורה עבור: " + name + " | תאריך עברי: " + hebDay + " " + hebMonth + " " + hebYear);
    
    // --- מחשבים גיל בסיסי ---
    // רושמים את שנת הלידה והשנה הנוכחית בגיליון log
    let age = null;
    let birthYear = null;
    let currentYear = null;
    let nextBirthdayYear = null;
    let useNextYearAge = false;
    
    try {
      let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
      if (logSheet) {
        // המרת שנת הלידה למספר
        birthYear = convertHebrewYearToNumber(hebYear);
        currentYear = getCurrentHebrewYear();
        
        // רושמים את הערכים בלוג
        logSheet.getRange(2, 1).setValue(birthYear);
        logSheet.getRange(2, 2).setValue(currentYear);
        
        // מחשבים את הגיל הבסיסי ורושמים בעמודה C
        age = calculateHebrewAge(birthYear, currentYear);
        logSheet.getRange(2, 3).setValue(age);
        
        Logger.log("[DEBUG] שנת לידה: " + birthYear + " | שנה נוכחית: " + currentYear + " | גיל בסיסי: " + age);
      }
    } catch (err) {
      Logger.log("[DEBUG] שגיאה בחישוב הגיל: " + err);
    }
    
    // --- ממירים לתאריך הלועזי הנכון ---
    let today = new Date(); // תאריך דינמי - התאריך הנוכחי
    Logger.log("[DEBUG] תאריך היום: " + today.toLocaleDateString());
    
    // המרת התאריך העברי ללועזי עבור השנה הנוכחית או הבאה
    let birthdayInfo = getNextBirthdayDateWithInfo(hebDay, hebMonth, today);
    if (!birthdayInfo || !birthdayInfo.date) {
      Logger.log("[DEBUG] לא הצלחנו להמיר את התאריך");
      continue;
    }
    
    const convertedDate = birthdayInfo.date;
    useNextYearAge = birthdayInfo.isNextYear;
    
    Logger.log("[DEBUG] התאריך הלועזי המתוכנן: " + convertedDate.toLocaleDateString());
    Logger.log("[DEBUG] האם משתמש בגיל של השנה הבאה? " + useNextYearAge);
    
    // מתקן את הגיל אם התאריך בשנה הבאה
    if (useNextYearAge && age !== null) {
      age = age + 1;
      Logger.log("[DEBUG] הגיל המתוקן (בשנה הבאה): " + age);
    }
    
    // רושמים את הגיל בעמודה E
    if (age !== null) {
      // בגלייונות אם מכניסים מספר שיכול להראות כמו תאריך, הוא מומר אוטומטית לתאריך
      // לכן נוסיף גרש לפני המספר כדי שיישאר כמספר
      sheet.getRange(i + 2, 5).setValue("'" + age);
      Logger.log("[DEBUG] רושם גיל בעמודה E: " + age);
    } else {
      sheet.getRange(i + 2, 5).setValue("");
    }
    
    // שומרים את התאריך המחושב בעמודה F
    sheet.getRange(i + 2, 6).setValue(convertedDate);
    Logger.log("[DEBUG] רושם תאריך בעמודה F: " + convertedDate.toLocaleDateString());
    
    // יוצרים או מעדכנים אירוע ביומן
    let newEventId = createOrUpdateCalendarEvent(
      name,
      convertedDate,
      eventId,
      `${hebDay} ${hebMonth}`,
      age  // מעבירים את הגיל המתוקן
    );
    
    if (newEventId && newEventId !== eventId) {
      // שומרים את מזהה האירוע בעמודה G
      sheet.getRange(i + 2, 7).setValue(newEventId);
      Logger.log("[DEBUG] רושם מזהה אירוע בעמודה G: " + newEventId);
    }
  }
}

/**
 * מחשב את התאריך הלועזי של יום ההולדת הבא עם מידע נוסף
 * מחזיר אובייקט עם התאריך והאם זו השנה הבאה
 */
function getNextBirthdayDateWithInfo(hebDay, hebMonth, today) {
  try {
    Logger.log("[FUNC] התחלת חישוב יום הולדת הבא: " + hebDay + " " + hebMonth);
    
    // תאריך דינמי אם לא מוגדר
    if (!today) {
      today = new Date(); // תאריך נוכחי
    }
    Logger.log("[FUNC] תאריך היום: " + today.toLocaleDateString());
    
    // מנקים את הנתונים
    hebDay = hebDay.toString().replace(/['"]/g, '');
    hebMonth = hebMonth.toString().replace(/['"]/g, '');
    
    // ממירים את היום למספר - בודקים אם זו אות עברית או מספר רגיל
    let day;
    if (!isNaN(hebDay)) {
      // אם זה כבר מספר - משתמשים בו
      day = parseInt(hebDay);
    } else {
      // אם זו אות עברית - ממירים למספר
      let hebrewLetterValues = {
        'א': 1, 'ב': 2, 'ג': 3, 'ד': 4, 'ה': 5, 'ו': 6, 'ז': 7, 'ח': 8, 'ט': 9,
        'י': 10, 'כ': 20, 'ל': 30
      };
      
      // בדיקה אם זה אות בודדת
      if (hebDay.length === 1) {
        day = hebrewLetterValues[hebDay];
        Logger.log("[FUNC] ממיר אות עברית " + hebDay + " למספר " + day);
      } else {
        // אם יש יותר מאות אחת, ננסה להמיר באמצעות פונקציית convertHebrewLettersToNumber
        day = convertHebrewLettersToNumber(hebDay);
        Logger.log("[FUNC] ממיר מספר עברי " + hebDay + " למספר " + day);
      }
    }
    
    if (isNaN(day) || day === null) {
      Logger.log("[FUNC] שגיאה: לא הצלחנו להמיר את היום למספר");
      return null;
    }
    Logger.log("[FUNC] היום לאחר המרה: " + day);
    
    // ממירים את החודש לפורמט Hebcal
    let monthMap = {
      'ניסן': 'Nisan', 'אייר': 'Iyar', 'סיון': 'Sivan',
      'תמוז': 'Tamuz', 'אב': 'Av', 'אלול': 'Elul',
      'תשרי': 'Tishrei', 'חשון': 'Cheshvan', 'כסלו': 'Kislev',
      'טבת': 'Tevet', 'שבט': 'Shvat', 'אדר': 'Adar'
    };
    
    let month = monthMap[hebMonth];
    if (!month) {
      Logger.log("[FUNC] שגיאה: החודש אינו תקין");
      return null;
    }
    
    // מקבלים את השנה העברית הנוכחית
    let currentHebrewYear = getCurrentHebrewYear();
    Logger.log("[FUNC] השנה העברית הנוכחית: " + currentHebrewYear);
    
    // המרת התאריך העברי ללועזי בשנה העברית הנוכחית
    let url = `https://www.hebcal.com/converter?cfg=json&hy=${currentHebrewYear}&hm=${month}&hd=${day}&h2g=1`;
    Logger.log("[FUNC] URL לשנה הנוכחית: " + url);
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    
    // יוצרים את התאריך מהתשובה של ה-API
    let dateThisYear = new Date(data.gy, data.gm - 1, data.gd);
    Logger.log("[FUNC] תאריך לשנה הנוכחית: " + dateThisYear.toLocaleDateString());
    
    // בדיקה אם התאריך כבר עבר השנה
    // נמיר את התאריכים למילישניות ללא שעות
    let todayWithoutTime = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    let dateThisYearWithoutTime = new Date(dateThisYear.getFullYear(), dateThisYear.getMonth(), dateThisYear.getDate()).getTime();
    
    Logger.log("[FUNC] מילישניות של היום: " + todayWithoutTime);
    Logger.log("[FUNC] מילישניות של התאריך בשנה הנוכחית: " + dateThisYearWithoutTime);
    Logger.log("[FUNC] האם התאריך כבר עבר? " + (dateThisYearWithoutTime < todayWithoutTime));
    
    // אם התאריך כבר עבר, נחשב לשנה הבאה
    if (dateThisYearWithoutTime < todayWithoutTime) {
      Logger.log("[FUNC] התאריך כבר עבר! מחשב לשנה הבאה...");
      
      // קוראים ל-API פעם נוספת עם השנה העברית הבאה
      url = `https://www.hebcal.com/converter?cfg=json&hy=${currentHebrewYear + 1}&hm=${month}&hd=${day}&h2g=1`;
      Logger.log("[FUNC] URL לשנה הבאה: " + url);
      response = UrlFetchApp.fetch(url);
      data = JSON.parse(response.getContentText());
      
      // יוצרים את התאריך החדש מהתשובה של ה-API
      let dateNextYear = new Date(data.gy, data.gm - 1, data.gd);
      Logger.log("[FUNC] תאריך לשנה הבאה: " + dateNextYear.toLocaleDateString());
      
      // מחזירים את התאריך ומידע נוסף
      return {
        date: dateNextYear,
        isNextYear: true
      };
    } else {
      Logger.log("[FUNC] התאריך עוד לא עבר, משתמש בתאריך השנה הנוכחית");
      // מחזירים את התאריך ומידע נוסף
      return {
        date: dateThisYear,
        isNextYear: false
      };
    }
  } catch (err) {
    Logger.log("[FUNC] שגיאה בחישוב יום הולדת הבא: " + err);
    return null;
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