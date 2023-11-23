const app = SpreadsheetApp.getActiveSpreadsheet();
const taobaoSheet = app.getSheetByName("taobao");
const storageSheet = app.getSheetByName("2023.11.23") // Change for new storage date
const ui = SpreadsheetApp.getUi();

function onOpen() {
  const menu = ui.createMenu('Check taobao');
  menu.addItem('Check in items', 'checkTaobaoItems').addToUi();

}

function checkTaobaoItems () {
  const data = taobaoSheet.getDataRange().getValues();
  const storageData = storageSheet.getDataRange().getValues();
  for(let i = 0; i < data.length; i++) {
    let tracking = data[i][1].toString();
    for(let j = 0; j < storageData.length; j++) {
      let storageTracking = storageData[j][0].toString();
      if(tracking === storageTracking) {
        Logger.log('Matched tracking: '+ tracking + ' -> storage: '+ storageTracking)
        checkedIn(i, j)
      } 
    }
  }

  // Check for Missing Items
  missingCheck();
}

function checkedIn( trackListPosition, storageListPosition) {
  // Logger.log(trackListPosition)
  // Logger.log(storageListPosition)
  let tbTrack = Math.floor(trackListPosition)
  let sgTrack = Math.floor(storageListPosition)
  taobaoSheet.getRange(`C${tbTrack + 1}`).setValue('Checked In').setFontColor('green')
  storageSheet.getRange(`B${sgTrack + 1}`).setValue('Checked In').setFontColor('green')
}

function missingCheck () {
  const data = taobaoSheet.getDataRange().getValues();
  for(let i = 0; i < data.length; i++) {
    let status = taobaoSheet.getRange(`C${i + 1}`).getValue();
    if(status === null || status === undefined || status === '') {
      taobaoSheet.getRange(`C${i + 1}`).setValue('Missing').setFontColor('red')
    }
  }

}
