const FILE_ID = '1NJh9mdp7SGaGL8-dOaYGm2DlOZ3rizsHY9VDTWQHxmA'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('UP GK Master Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      // This tells Google NOT to force a login check for this UI
      .setSandboxMode(HtmlService.SandboxMode.IFRAME) 
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getChapters() {
  const ss = SpreadsheetApp.openById(FILE_ID); 
  const sheet = ss.getSheetByName("Sheet2"); 
  
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  const data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
  return [...new Set(data.flat())].map(c => c.toString().trim()).filter(c => c.length > 0).sort();
}

/** * FAST LIST: Returns only basic info to build the navigation skeleton 
 */
function getQuestionList(chapterName) {
  const ss = SpreadsheetApp.openById(FILE_ID);
  const sheet = ss.getSheetByName("Sheet2"); 
  
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const target = chapterName.toString().trim();
  
  return data.map((row, index) => ({
    rowIdx: index,
    chapter: row[0],
    question: row[1]
  })).filter(item => item.chapter === target && item.question !== "");
}

/** * DETAIL FETCH: Fetches the options/explanation for a specific row 
 */
function getQuestionDetails(rowIdx) {
  const ss = SpreadsheetApp.openById(FILE_ID);
  const sheet = ss.getSheetByName("Sheet2"); 
  
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  const row = sheet.getRange(rowIdx + 1, 1, 1, 10).getValues()[0];
  
  return {
    Question: row[1], OptionA: row[2], OptionB: row[3], OptionC: row[4], 
    OptionD: row[5], OptionE: row[6], Answer: row[7], Explanation: row[8], Exam: row[9]
  };
}

/** * BACKGROUND SYNC: Fetches all full data for the chapter 
 */
function getFullChapterData(chapterName) {
  const ss = SpreadsheetApp.openById(FILE_ID); 
  const sheet = ss.getSheetByName("Sheet2"); 
  
  if (!sheet) {
    Logger.log("Sheet not found!");
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const target = chapterName.toString().trim();

  return data.filter(row => row[0] && row[0].toString().trim() === target && row[1])
    .map((row, index) => {
      let obj = {};
      headers.forEach((header, i) => { obj[header.toString().replace(/\s+/g, "")] = row[i]; });
      return obj;
    });
}