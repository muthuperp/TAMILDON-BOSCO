function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('à®¤à®®à®¿à®´à¯ à®µà®¿à®©à®¾à®Ÿà®¿ à®µà®¿à®©à®¾')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getQuestions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Questions");
  if(!sheet) throw new Error("Questions sheet not found!");
  const data = sheet.getDataRange().getValues();
  let questions = [];
  for(let i=1; i<data.length; i++) {
    let row = data[i];
    if(row[0] && row[1]) {
      questions.push({
        question: row[0],
        options: [row[1], row[2], row[3], row[4]],
        answer: parseInt(row[5]) - 1
      });
    }
  }
  return questions.sort(() => Math.random() - 0.5);
}

function submitQuizResult(name, rollNo, school, email, score, total, timeTaken) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  if(!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Results");
    sheet.appendRow(["Timestamp", "Name", "Roll No", "School", "Email", "Score", "Total", "Time Taken (s)"]);
  }
  sheet.appendRow([new Date(), name, rollNo, school, email, score, total, timeTaken]);
}

function getLeaderboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return []; 
  
  let students = [];
  for(let i=1; i<data.length; i++) {
    let row = data[i];
    if(row[1]) {
      students.push({
        name: row[1],
        score: parseInt(row[5]) || 0,
        time: parseInt(row[7]) || 9999
      });
    }
  }
  students.sort((a, b) => {
    if(b.score !== a.score) return b.score - a.score;
    return a.time - b.time;
  });
  return students.slice(0, 10);
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
