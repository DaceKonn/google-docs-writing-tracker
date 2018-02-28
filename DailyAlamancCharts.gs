function buildWordsWrittenChart() {
  var codeName = "wordsWritten";
  var sheet = SpreadsheetApp.openById(WRITING_DATA).getSheetByName(WRITING_SHEET);
  var chartBuilder = sheet.newChart();
  var lastRow = sheet.getLastRow();
  var minRow = Math.max(lastRow - 365, 1);
  
  var chart = chartBuilder
  .addRange(sheet.getRange(COL_DATES + minRow + ":" + COL_DATES + lastRow))
  .addRange(sheet.getRange(COL_WRITING_TOTAL + minRow + ":" + COL_WRITING_TOTAL + lastRow))
  .addRange(sheet.getRange(COL_AVERAGE + minRow + ":" + COL_AVERAGE + lastRow))
  .addRange(sheet.getRange(COL_GOAL + minRow + ":" + COL_GOAL + lastRow))
  .asLineChart()
  .setOption('title', 'Words Written')
  .setXAxisTitle('Time')
  .setXAxisTitle('Words')
  .setRange(0, 300)
  .build();
  
  
  return { 
    message: "<img src='cid:"+codeName+"'><br>", 
    image: chart.getAs("image/png").setName(codeName),
    'codeName': codeName
  };
}

function buildWordCountChart() {
  var codeName = "wordCount";
  var sheet = SpreadsheetApp.openById(WRITING_DATA).getSheetByName(WRITING_SHEET);
  var chartBuilder = sheet.newChart();
  var lastRow = sheet.getLastRow();
  var minRow = Math.max(lastRow - 365, 1);
  
  var chart = chartBuilder
  .addRange(sheet.getRange(COL_DATES + minRow + ":" + COL_DATES + lastRow))
  .addRange(sheet.getRange(COL_WRITING_TOTAL + minRow + ":" + COL_WRITING_TOTAL + lastRow))
  .addRange(sheet.getRange(COL_AVERAGE + minRow + ":" + COL_AVERAGE + lastRow))
  .addRange(sheet.getRange(COL_GOAL + minRow + ":" + COL_GOAL + lastRow))
  .asAreaChart()
  .setOption('title', 'Word Count')
  .setXAxisTitle('Days')
  .setYAxisTitle('Words')
  .setRange(0, 300)
  .build();
  
  
  return { 
    message: "<img src='cid:"+codeName+"'><br>", 
    image: chart.getAs("image/png").setName(codeName),
    'codeName': codeName
  };
}

function buildWritingTimeChart() {
  var codeName = "writingTime";
  var sheet = SpreadsheetApp.openById(WRITING_DATA).getSheetByName(WRITING_SHEET);
  var chartBuilder = sheet.newChart();
  var lastRow = sheet.getLastRow();
  var minRow = Math.max(lastRow - 365, 1);
  
  var chart = chartBuilder
  .addRange(sheet.getRange(COL_DATES + minRow + ":" + COL_DATES + lastRow))
  .addRange(sheet.getRange(COL_WRITING_TIME + minRow + ":" + COL_WRITING_TIME + lastRow))
  .asAreaChart()
  .setOption('title', 'Writing Time')
  .setXAxisTitle('Days')
  .setYAxisTitle('min')
  .setRange(0, 120)
  .build();
  
  
  return { 
    message: "<img src='cid:"+codeName+"'><br>", 
    image: chart.getAs("image/png").setName(codeName),
    'codeName': codeName
  };
}

function buildWordsPerMinuteChart() {
  var codeName = "wordsPerMinute";
  var sheet = SpreadsheetApp.openById(WRITING_DATA).getSheetByName(WRITING_SHEET);
  var chartBuilder = sheet.newChart();
  var lastRow = sheet.getLastRow();
  var minRow = Math.max(lastRow - 365, 1);
  
  var chart = chartBuilder
  .addRange(sheet.getRange(COL_DATES + minRow + ":" + COL_DATES + lastRow))
  .addRange(sheet.getRange(COL_WRITING_WORDS_MINUTE + minRow + ":" + COL_WRITING_WORDS_MINUTE + lastRow))
  .asAreaChart()
  .setOption('title', 'Words Per Minute')
  .setXAxisTitle('Days')
  .setYAxisTitle('Words')
  .setRange(0, 30)
  .build();
  
  
  return { 
    message: "<img src='cid:"+codeName+"'><br>", 
    image: chart.getAs("image/png").setName(codeName),
    'codeName': codeName
  };
}

function buildTimeAndWordsChart() {
  var codeName = "timeAndWords";
  var sheet = SpreadsheetApp.openById(WRITING_DATA).getSheetByName(WRITING_SHEET);
  var chartBuilder = sheet.newChart();
  var lastRow = sheet.getLastRow();
  var minRow = Math.max(lastRow - 365, 1);
  
  var chart = chartBuilder
  .addRange(sheet.getRange(COL_DATES + minRow + ":" + COL_DATES + lastRow))
  .addRange(sheet.getRange(COL_WRITING_TIME + minRow + ":" + COL_WRITING_TIME + lastRow))
  .addRange(sheet.getRange(COL_WRITING_TOTAL + minRow + ":" + COL_WRITING_TOTAL + lastRow))
  .asColumnChart()
  .setOption('title', 'Time and Words')
  .setOption('vAxes', {1: {title:'a',textStyle: {color: 'red'}}})
  .setOption('series', {2: {targetAxisIndex:1}})

  .setXAxisTitle('Days')
  //.setYAxisTitle('Words')
  .setRange(0, 30)
  .build();
  
  
  return { 
    message: "<img src='cid:"+codeName+"'><br>", 
    image: chart.getAs("image/png").setName(codeName),
    'codeName': codeName
  };
}