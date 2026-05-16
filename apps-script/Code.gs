function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Health Tracker')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}



/**
 * Get weekday from date
 */
function getWeekday(dateString) {

  const date = new Date(dateString);

  return Utilities.formatDate(

    date,

    Session.getScriptTimeZone(),

    'EEEE'
  );
}



/**
 * Load checklist tasks
 */
function getChecklist(
  period,
  selectedDate
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  let sheetName = '';



  if (period === 'morning') {

    sheetName =
      'WAKE FLOW / MORNING RITUAL';
  }

  else if (period === 'midday') {

    sheetName =
      'MIDDAY BLOCK';
  }

  else if (period === 'evening') {

    sheetName =
      'WIND-DOWN / SLEEP ONBOARDING';
  }



  const weekday =

    getWeekday(selectedDate)
      .toLowerCase();



  const checklistSheet =
    ss.getSheetByName(sheetName);



  const tasksSheet =
    ss.getSheetByName('Tasks');



  const checklistData =
    checklistSheet
      .getDataRange()
      .getValues();



  const tasksData =
    tasksSheet
      .getDataRange()
      .getValues();



  /**
   * Build task map
   */
  const taskMap = {};



  for (let i = 1; i < tasksData.length; i++) {

    const taskId =

      tasksData[i][0]
        .toString()
        .trim();



    const taskText =
      tasksData[i][1];



    taskMap[taskId] = taskText;
  }



  const items = [];



  for (let i = 1; i < checklistData.length; i++) {

    const rowWeekday =

      checklistData[i][0]
        .toString()
        .trim()
        .toLowerCase();



    const taskId =

      checklistData[i][1]
        .toString()
        .trim();



    if (rowWeekday === weekday) {

      items.push({

        id: taskId,

        text:
          taskMap[taskId] || taskId

      });
    }
  }



  return items;
}



/**
 * Prevent duplicate submissions
 */
function checkPeriodSubmission(
  email,
  selectedDate,
  period
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  const responseSheet =
    ss.getSheetByName('Response');



  const data =
    responseSheet
      .getDataRange()
      .getValues();



  const headers = data[0];



  let taskIds = [];



  /**
   * Determine period columns
   */
  if (period === 'morning') {

    taskIds = [

      'M1','M2','M3','M4','M5','M6',

      'M7','M8','M9','M10','M11','M12'

    ];
  }

  else if (period === 'midday') {

    taskIds = [

      'B1','B2','B3','B4','B5','B6',

      'B7','B8','B9','B10','B11','B12'

    ];
  }

  else if (period === 'evening') {

    taskIds = [

      'W1','W2','W3','W4','W5','W6',

      'W7','W8','W9','W10','W11'

    ];
  }



  /**
   * Search existing rows
   */
  for (let i = 1; i < data.length; i++) {

    const rowEmail =

      data[i][2]
        .toString()
        .trim()
        .toLowerCase();



    const rowDate =

      Utilities.formatDate(

        new Date(data[i][3]),

        Session.getScriptTimeZone(),

        'yyyy-MM-dd'
      );



    if (

      rowEmail ===
      email
        .trim()
        .toLowerCase()

      &&

      rowDate === selectedDate

    ) {

      for (let j = 0; j < taskIds.length; j++) {

        const columnIndex =
          headers.indexOf(taskIds[j]);



        if (

          columnIndex !== -1

          &&

          data[i][columnIndex] !== ''

        ) {

          return true;
        }
      }
    }
  }



  return false;
}



/**
 * Save checklist
 */
function submitChecklist(formData) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  const responseSheet =
    ss.getSheetByName('Response');



  const data =
    responseSheet
      .getDataRange()
      .getValues();



  const headers =

    data[0].map(h =>
      h.toString().trim()
    );



  const timestamp =
    new Date();



  const selectedDate =
    formData.selectedDate;



  const weekday =
    getWeekday(selectedDate);



  let existingRow = -1;



  /**
   * Search for existing row
   */
  for (let i = 1; i < data.length; i++) {

    const rowEmail =

      data[i][2]
        .toString()
        .trim()
        .toLowerCase();



    const rowDate =

      Utilities.formatDate(

        new Date(data[i][3]),

        Session.getScriptTimeZone(),

        'yyyy-MM-dd'
      );



    if (

      rowEmail ===

      formData.email
        .trim()
        .toLowerCase()

      &&

      rowDate === selectedDate

    ) {

      existingRow = i + 1;

      break;
    }
  }



  /**
   * Create row if not exists
   */
  if (existingRow === -1) {

    const emptyRow =

      new Array(headers.length)
        .fill('');



    emptyRow[0] = timestamp;
    emptyRow[1] = formData.name;
    emptyRow[2] = formData.email;
    emptyRow[3] = selectedDate;
    emptyRow[4] = weekday;



    responseSheet.appendRow(emptyRow);



    existingRow =
      responseSheet.getLastRow();
  }



  /**
   * Save task results
   */
  formData.items.forEach(item => {

    const columnIndex =
      headers.indexOf(item.id);



    if (columnIndex !== -1) {

      responseSheet

        .getRange(
          existingRow,
          columnIndex + 1
        )

        .setValue(

          item.completed
            ? 'Yes'
            : 'No'
        );
    }

  });



  return {

    success: true,

    message:
      'Checklist submitted successfully!'
  };
}
