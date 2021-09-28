//creates and populates a Google Form with random questions from the test Bank

//specifics
const finalRow = 489; //row of the LAST QUESTION IN THE BANK
let quantity = 100; //how many questions in the test
let formRow=22;

function makeTests() {
  let startRow = 3;
  let endRow = 3;
  for (let i = startRow; i <= endRow; i++) {
    //createNewTest(i);
  }
}

function createNewTest() {
  const MASTER_TEST_SHEET = SpreadsheetApp.openById(**********).getSheetByName("Tests");
  let form_info = MASTER_TEST_SHEET.getRange(formRow, 1, 1, 2).getValues();
  console.log(form_info);
  let time = form_info[0][1];
  let testCode = form_info[0][0];

  //constants
  const CONFIRMATION = "Congratulations on completing the exam for regionals! We will announce the top 3 winners during the virtual conference and release individual scores via email afterwards. In the meantime, if you have any questions, please email us at officialmarylanddeca@gmail.com.";
  const FORM_TITLE = "2021 Regionals Exam";
  const CLOSED_MESSAGE = "The DECA Regionals Exam will only be open during the specified scheduled intervals. Please email ******* for assistance."
  const DESCRIPTION = "When filling out your email address, please use your personal email from the regionals registration form. REMEMBER: you are also responsible for hitting the submit button by " + time + " in order to receive a score. If you encounter a technical issue that prevents you from completing the exam, please email ************.";

  const MASTER_FOLDER = DriveApp.getFolderById(**************);

  const document = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Questions');

  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Test " + testCode);
  let titleRange = outputSheet.getRange(1, 1, 1, 3);
  titleRange.setValues([
    ["Question Number", "Question", "Source"]
  ])
  titleRange.setFontWeight("bold");
  outputSheet.setFrozenRows(1);

  const imgs = MASTER_FOLDER.createFolder("IMG_Q " + testCode);

  const form = FormApp.create("Regionals Test Form " + testCode);
  DriveApp.getFileById(form.getId()).moveTo(MASTER_FOLDER);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ****************);
  form.setAcceptingResponses(false);
  form.setDescription(DESCRIPTION).setConfirmationMessage(CONFIRMATION);
  form.setIsQuiz(true);
  form.setShuffleQuestions(true);
  form.setCollectEmail(true);
  form.setTitle(FORM_TITLE);
  form.setShowLinkToRespondAgain(false).setAllowResponseEdits(false);
  form.setCustomClosedFormMessage(CLOSED_MESSAGE)

  form.addTextItem().setTitle("Full Name").setRequired(true);
  form.addPageBreakItem();

  console.log("Made " + form.getTitle() + " at " + form.getEditUrl());
  MASTER_TEST_SHEET.getRange(formRow, 3, 1, 4).setValues([[imgs.getUrl(), form.getEditUrl(), form.getPublishedUrl(), "TRUE"]]);

  let questionIndexes = getNRandomNumbers(2, finalRow, quantity);
  for (let i = 0; i < questionIndexes.length; i++) {
    let qIndex = questionIndexes[i];
    questionInfo = document.getRange(qIndex, 1, 1, 7).getValues()[0];

    try {
    Utilities.sleep(100);
    let image = textToImage(questionInfo[0], i + 1, imgs, testCode);
    console.log("Created image " + image.getName() + " at " + image.getUrl());

    let answer = questionInfo[5].charCodeAt(0) - "A".charCodeAt(0);
    let item = form.addMultipleChoiceItem();
    item.setRequired(false).setPoints(1);
    //item.setTitle(questionInfo[0]);
    item.setChoices([
      item.createChoice(questionInfo[1], 0 == answer),
      item.createChoice(questionInfo[2], 1 == answer),
      item.createChoice(questionInfo[3], 2 == answer),
      item.createChoice(questionInfo[4], 3 == answer),
    ]);
    console.log("Created question from index " + qIndex + ": " + questionInfo[0])

    outputSheet.getRange(i + 2, 1, 1, 3).setValues([
      [qIndex, questionInfo[0], questionInfo[6]]
    ]);
  } catch (e) {
    console.log("error: " + e + " ... trying again")
    i--;
  }
  }
}


//helpers
function textToImage(question, number, folder, form) {
  data = {
    "question": question
  };
  options = {
    "method": "post",
    "payload": data
  };
  let dataURI = UrlFetchApp.fetch(************, options).getContentText()
  var type = (dataURI.split(";")[0]).replace('data:', '');
  var imageUpload = Utilities.base64Decode(dataURI.split(",")[1]);
  var blob = Utilities.newBlob(imageUpload, type, "Form " + form + " Question " + number + "." + type.replace('image/', ''));
  return folder.createFile(blob);
}

function getRandomNumber(min, max) {
  return Math.random() * (max - min) + min;
}

function getNRandomNumbers(from, to, n) {
  var listNumbers = [];
  var nRandomNumbers = [];
  for (let i = from; i <= to; i++) {
    listNumbers.push(i);
  }
  for (let i = 0; i < n; i++) {
    var index = getRandomNumber(0, listNumbers.length);
    nRandomNumbers.push([listNumbers[parseInt(index)]]);
    listNumbers.splice(index, 1);
  }
  return nRandomNumbers;
}
