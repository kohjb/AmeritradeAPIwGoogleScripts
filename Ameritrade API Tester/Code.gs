//Script to test all the functions in the Ameritrade API Library (amtd)

function onOpen() {  
  //create menu items
  var currentssheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Authenticate', functionName: 'amtd_showPane'}
  ];
  
  //Add the menu to the sheet
  currentssheet.addMenu("Stock Portfolio", menuItems);
}



