//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//STORAGE ARRAYS
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//change the value of var SEARCH_QUERY = "  " below, to match your desired label;
var SEARCH_QUERY = " ";
var subjects = [];
var bodys = [];
var dates = [];
var froms = [];
var names = [];
var emailPath = [];
var tos = []; //Cristian added

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//RUNS THE ENTIRE SCRIPT WHEN THE DOCUMENT IS OPENED
//ADDITIONALLY, YOU CAN ADD TIME BASED TRIGGERS FROM THE GOOGLE SCRIPT PROJECTS DASHBOARD
//https://script.google.com/home
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function onOpen() {
  SpreadsheetApp.getActiveSheet()//.clear();
  addHeader(SpreadsheetApp.getActiveSheet());
  getInfo_(SEARCH_QUERY);
  
  if (this.froms.length != 0) {
    appendData_(SpreadsheetApp.getActiveSheet(), this.froms, this.tos, this.subjects, this.bodys, this.names, this.dates, this.emailPath);
    } 
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//ADDS HEADER INFO - IF THERE IS NO HEADER
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function addHeader(Sheet1){
  Sheet1.setRowHeight(1, 21)
  var i = 1;
  var headers = ['From','To','Subject','Body','Name','Date','URL Reply'];
  
  if(!Sheet1.getLastRow()){
    while( i <= headers.length){
    if( i === 1){
        Sheet1.getRange(Sheet1.getLastRow()+1, i).setValue(headers[0]);
        Sheet1.getRange(Sheet1.getLastRow(), i).setBackground('white');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontColor('black');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontSize(10);
        i++;
    }else {
        Sheet1.getRange(Sheet1.getLastRow(), i).setValue(headers[i-1]);
        Sheet1.getRange(Sheet1.getLastRow(), i).setBackground('white');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontColor('black');
        Sheet1.getRange(Sheet1.getLastRow(), i).setFontSize(10);
        i++;
         }
    };
  };
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//GETS INFO FROM EMAIL AND STORES IN ARRAYS
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function getInfo_(q) {
var label = GmailApp.getUserLabelByName(q); //Replaced "GmailApp.getInboxThreads(0,500);" with  "GmailApp.getUserLabelByName("Astrivis CRM");" //
var threads = label.getThreads();
for (var i = 0; i < threads.length; i++)  {
      var msgs = threads[i].getMessages();
      for (var j in msgs) {          
          var message = msgs[j];
          this.bodys.push([message.getPlainBody().replace(/<.*?>/g, '').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
          this.subjects.push([message.getSubject().replace(/<.*?>/g, '\n').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
          this.dates.push([message.getDate()]);
          this.names.push([message.getFrom().replace(/<.*?>/g, '').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
          this.froms.push([message.getFrom().replace(/^.+</, '').replace(">", '') ]);
          this.tos.push([message.getTo().replace(/^.+</, '').replace(">", '') ]);  //Cristian added
          this.emailPath.push(['https://mail.google.com/mail/u/0/#inbox/' + message.getId()]);
    }
  }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//APPENDS ARRAY INFO TO THE SPREADSHEET
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
function appendData_(Sheet1, array2dFroms, array2dtos, array2dSubjects, array2dBodys, array2dNames, array2dDates, array2dURL) {
  Sheet1.getRange(Sheet1.getLastRow() + 1, 1, array2dFroms.length, array2dFroms[0].length).setValues(array2dFroms);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 2, array2dtos.length, array2dtos[0].length).setValues(array2dtos);   //Cristian added
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 3, array2dSubjects.length, array2dSubjects[0].length).setValues(array2dSubjects);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 4, array2dBodys.length, array2dBodys[0].length).setValues(array2dBodys);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 5, array2dNames.length, array2dNames[0].length).setValues(array2dNames);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 6, array2dDates.length, array2dDates[0].length).setValues(array2dDates);
  Sheet1.getRange(Sheet1.getLastRow() + 1 - array2dFroms.length, 7, array2dURL.length, array2dURL[0].length).setValues(array2dURL);
    }
