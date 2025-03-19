/**
 * ממפה את שם החודש העברי לפורמט הנדרש ב-API של Hebcal.
 * חשוב במיוחד לטיפול ב'אדר א'/'אדר ב'.
 *
 * @param {string} hebMonth - שם החודש העברי
 * @return {string} - שם החודש באנגלית בפורמט המתאים ל-Hebcal
 */
function mapHebrewMonthToHebcal(hebMonth) {
  if (!hebMonth) {
    Logger.log("שגיאה: חודש עברי הוא undefined או null");
    return "";
  }
  
  Logger.log("ממפה חודש עברי: " + hebMonth);
  
  switch (hebMonth.trim()) {
    case "תשרי":   return "Tishrei";
    case "חשון":   return "Cheshvan";
    case "כסלו":   return "Kislev";
    case "טבת":    return "Tevet";
    case "שבט":    return "Shevat";
    case "אדר א":  return "Adar1";  // בשנה מעוברת
    case "אדר ב":  return "Adar2";  // בשנה מעוברת
    case "אדר":    return "Adar";   // בשנה רגילה
    case "ניסן":   return "Nisan";
    case "אייר":   return "Iyar";
    case "סיון":   return "Sivan";
    case "תמוז":   return "Tammuz";
    case "אב":     return "Av";
    case "אלול":   return "Elul";
    default:
      Logger.log("שגיאה: חודש עברי לא מזוהה: " + hebMonth);
      return ""; // חודש לא מזוהה
  }
}

/**
 * ממיר שנה עברית מפורמט טקסט (למשל תשפ"ג) למספר (5783)
 * @param {string} hebrewYear - השנה בפורמט עברי
 * @return {number} - השנה כמספר
 */
function parseHebrewYear(hebrewYear) {
  if (!hebrewYear) return getCurrentHebrewYear();
  
  // אם זה כבר מספר
  if (!isNaN(hebrewYear)) {
    return parseInt(hebrewYear, 10);
  }

  // מפה של אותיות לערכים
  const letterValues = {
    'א': 1, 'ב': 2, 'ג': 3, 'ד': 4, 'ה': 5,
    'ו': 6, 'ז': 7, 'ח': 8, 'ט': 9, 'י': 10,
    'כ': 20, 'ל': 30, 'מ': 40, 'נ': 50, 'ס': 60,
    'ע': 70, 'פ': 80, 'צ': 90, 'ק': 100, 'ר': 200,
    'ש': 300, 'ת': 400
  };

  // מנקים את הטקסט מגרשיים וסימנים מיוחדים
  let cleanYear = hebrewYear.replace(/['"״]/g, '');
  Logger.log("שנה אחרי ניקוי: " + cleanYear); // תשפג
  
  // מחלקים לאלפים (תש) ומאות (פג)
  let prefix = cleanYear.substring(0, 2); // תש
  let suffix = cleanYear.substring(2);    // פג
  
  Logger.log("תחילית השנה: " + prefix); // תש
  Logger.log("סיומת השנה: " + suffix);  // פג
  
  // מחשבים את ערך התחילית (תש)
  let prefixValue = 0;
  for (let char of prefix) {
    prefixValue += letterValues[char] || 0;
  }
  Logger.log("ערך התחילית: " + prefixValue); // 700
  
  // מחשבים את ערך הסיומת (פג)
  let suffixValue = 0;
  for (let char of suffix) {
    suffixValue += letterValues[char] || 0;
  }
  Logger.log("ערך הסיומת: " + suffixValue); // 83
  
  // השנה המלאה היא 5000 + ערך התחילית + ערך הסיומת
  let fullYear = 5000 + prefixValue + suffixValue;
  Logger.log("השנה המלאה: " + fullYear); // 5783
  
  return fullYear;
}

/**
 * פונקציה להמרת תאריך עברי (יום, חודש, שנה) לתאריך לועזי (אובייקט Date).
 * משתמשים ב-API של Hebcal.
 *
 * @param {string|number} hebDay - יום בחודש העברי (למשל "כ"ג" או "23")
 * @param {string} hebMonth - שם החודש העברי (למשל "אדר א", "תשרי", וכו')
 * @param {string|number} [hebYear] - שנה עברית (אופציונלי)
 * @param {boolean} [checkNextYear=true] - האם לבדוק את השנה הבאה אם התאריך עבר
 * @return {Date|null} - תאריך לועזי כאובייקט Date, או null אם לא מצליחים להמיר
 */
function convertHebrewToGregorian(hebDay, hebMonth, hebYear, checkNextYear = true) {
  try {
    Logger.log("מתחיל המרת תאריך עברי: יום=" + hebDay + ", חודש=" + hebMonth + ", שנה=" + hebYear);
    
    // ממירים את היום למספר
    let dayNumber = parseHebrewDay(hebDay); 
    Logger.log("יום עברי אחרי המרה למספר: " + dayNumber);
    
    if (!dayNumber) {
      Logger.log("שגיאה: לא הצלחנו להמיר את היום העברי למספר");
      return null;
    }

    let mappedMonth = mapHebrewMonthToHebcal(hebMonth);
    Logger.log("חודש עברי אחרי המרה לפורמט Hebcal: " + mappedMonth);
    
    if (!mappedMonth) {
      Logger.log("שגיאה: לא הצלחנו למפות את החודש העברי");
      return null;
    }

    // המרת השנה העברית למספר
    let hy = parseHebrewYear(hebYear);
    Logger.log("שנה עברית אחרי המרה למספר: " + hy);

    // בניית ה-URL ל-API
    let url = "https://www.hebcal.com/converter?cfg=json"
            + "&hy=" + hy
            + "&hm=" + mappedMonth
            + "&hd=" + dayNumber
            + "&h2g=1"; // המרה מעברי ללועזי

    Logger.log("URL ל-API: " + url);

    let response = UrlFetchApp.fetch(url);
    let json = JSON.parse(response.getContentText());
    Logger.log("תשובה מה-API: " + JSON.stringify(json));

    // Hebcal מחזיר gy (שנה לועזית), gm (חודש לועזי), gd (יום לועזי)
    let gy = json.gy;
    let gm = json.gm;
    let gd = json.gd;
    
    Logger.log("תאריך לועזי שהתקבל: " + gd + "/" + gm + "/" + gy);

    if (!gy || !gm || !gd) {
      Logger.log("שגיאה: חסרים נתונים בתשובה מה-API");
      return null;
    }

    // יוצרים אובייקט Date
    let convertedDate = new Date(gy, gm - 1, gd);
    
    // בודקים אם צריך לחשב את השנה הבאה
    if (checkNextYear) {
      const today = new Date();
      if (convertedDate < today) {
        Logger.log("התאריך כבר עבר השנה, מנסה להמיר לשנה הבאה");
        const currentHebrewYear = getCurrentHebrewYear();
        return convertHebrewToGregorian(dayNumber, hebMonth, currentHebrewYear, false);
      }
    }
    
    return convertedDate;

  } catch (err) {
    Logger.log("שגיאה בהמרת תאריך: " + err);
    return null;
  }
}

/**
 * פונקציה פשוטה שמחזירה את השנה העברית הנוכחית.
 * משתמשים ב-API של Hebcal כדי לקבל את השנה העברית הנוכחית.
 *
 * @return {number} - השנה העברית הנוכחית
 */
function getCurrentHebrewYear() {
  try {
    const today = new Date();
    const gy = today.getFullYear();
    const gm = today.getMonth() + 1;
    const gd = today.getDate();

    let url = "https://www.hebcal.com/converter?cfg=json"
            + "&gy=" + gy
            + "&gm=" + gm
            + "&gd=" + gd
            + "&g2h=1"; // המרה מלועזי לעברי

    Logger.log("מבקש את השנה העברית הנוכחית מ-Hebcal: " + url);
    
    let response = UrlFetchApp.fetch(url);
    let json = JSON.parse(response.getContentText());
    Logger.log("תשובה מה-API לשנה נוכחית: " + JSON.stringify(json));
    
    // Hebcal מחזיר את השנה העברית בשדה hy
    return json.hy;
  } catch (err) {
    Logger.log("שגיאה בקבלת שנה עברית נוכחית: " + err);
    return 5784; // ברירת מחדל - תשפ"ד
  }
}

/**
 * פונקציה שעוזרת להמיר תווים כמו "כ"ג" למספר 23.
 * אם מקבל מספר רגיל (23), פשוט מחזירה אותו כ-Number.
 *
 * @param {string|number} dayInput - היום בעברית או במספר
 * @return {number} - היום כמספר
 */
function parseHebrewDay(dayInput) {
  // אם כבר מספר
  if (!isNaN(dayInput)) {
    return parseInt(dayInput, 10);
  }
  // אפשרות: מפה של אותיות עבריות למספרים, או להשתמש בספרייה חיצונית.
  // כאן לצורך פשטות ננסה רק "קצת" פונקציונליות:
  // לדוגמה "י"ב" -> 12, "כ"ג" -> 23, "ט"ו" -> 15, וכו'.
  // אפשרויות מתקדמות ידרשו מימוש מלא של גימטריה.

  const gimatriaMap = {
    'א':1, 'ב':2, 'ג':3, 'ד':4, 'ה':5, 'ו':6, 'ז':7, 'ח':8, 'ט':9,
    'י':10, 'כ':20, 'ל':30, 'מ':40, 'נ':50, 'ס':60, 'ע':70, 'פ':80, 'צ':90,
    'ק':100, 'ר':200, 'ש':300, 'ת':400
  };

  // מחלקים את המחרוזת לסימנים ומחברים
  // לדוגמה "כ"ג" -> ['כ','ג']
  let chars = dayInput.replace("״","").replace("''","").split(/['"]/); 
  // לעתים נוסיף טיפול להורדת גרשיים. ננקה גם רווחים
  let cleanStr = dayInput.replace(/[^\u0590-\u05FF]/g, '').split('');
  let sum = 0;
  cleanStr.forEach(ch => {
    if (gimatriaMap[ch]) {
      sum += gimatriaMap[ch];
    }
  });

  // אם לא הצלחנו, נחזיר 0
  return sum > 0 ? sum : 0;
}
