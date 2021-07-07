const sheet_config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
const sheet_hidden = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('-hidden');
const sheet_accountStructure = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('accountStructure');
const sheet_customDimensions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('customDimensions');
const sheet_customDimensionsMatrix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('customDimensionsMatrix');
const sheet_customDimensionsMatrixPivot = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('-customDimensionsMatrixPivot');
const sheet_filters = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('filters');
const sheet_viewFilters = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('viewFilters');
const sheet_viewFiltersMatrix = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('viewFiltersMatrix');
const sheet_viewFiltersMatrixPivot = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('-viewFiltersMatrixPivot');
const sheet_goals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('goals');

let data = [];

// ---------- Build menu options ----------
function onOpen() {
  clearSheets();
  sheet_config.getRange('C2:C2').setValue('Select an account...');
  sheet_config.getRange('C4:C4').setValue('all');

  SpreadsheetApp.getUi().createMenu('SDR Builder')
  .addItem('Clear Data','clearSheets')
  .addItem('Get Accounts','getAccounts')
  .addItem('Get Account Structure','getAccountStructure')
  .addItem('Get Custom Dimensions','getCustomDimensions')
  .addItem('Get Filters','getFilters')
  .addItem('Get Goals','getGoals')
  .addToUi();
}


// ---------- Clears all data from visible sheets ----------
function clearSheets() {
  sheet_accountStructure.setTabColor(null);
  sheet_accountStructure.clearContents();
  sheet_hidden.getRange(2,1,1000,1).clearContent();
  sheet_hidden.getRange(3,3,1000,1).clearContent();
  sheet_customDimensions.setTabColor(null);
  sheet_customDimensions.clearContents();
  sheet_customDimensionsMatrix.setTabColor(null);
  sheet_filters.clearContents();
  sheet_filters.setTabColor(null);
  sheet_viewFilters.clearContents();
  sheet_viewFilters.setTabColor(null);
  sheet_viewFiltersMatrixPivot.getRange(1,1,sheet_viewFiltersMatrixPivot.getLastRow(),4).clearContent();
  sheet_viewFiltersMatrixPivot.setTabColor(null);
  sheet_viewFiltersMatrix.setTabColor(null);
  sheet_goals.clearContents();
  sheet_goals.setTabColor(null);
}


// ---------- Writes a row of data to a given sheet ----------
function writeData(sheet,output,hidden) {
  sheet.getRange(1,1,output.length,output[0].length).clearContent();
  sheet.getRange(1,1,output.length,output[0].length).setValues(output);
  sheet.autoResizeColumns(1,output[0].length);
  sheet.getRange(1,1,output.length,output[0].length).setHorizontalAlignment('left');

  if (hidden !== "yes") {
    sheet.setActiveRange(sheet.getRange(sheet.getLastRow() + 1,1));
    sheet.activate();
    sheet.setTabColor('#0000FF');
  }
}


// ---------- Output available accounts for the dropdown selector ----------
function getAccounts() {
  sheet_hidden.getRange(3,3,1000,1).clearContent();
  sheet_hidden.getRange(2,1,1000,1).clearContent();

  let responseAcc = Analytics.Management.Accounts.list();
  
  for (let i = 0; i < responseAcc['totalResults']; i++) {
    let accountId = responseAcc['items'][i]['id'];
    let accountName = responseAcc['items'][i]['name'];

    data.push([`${accountId} | ${accountName}`]);
  }

  sheet_hidden.getRange(2,1,data.length,data[0].length).setValues(data);
}


// ---------- Output account, property, and view details ----------
function getAccountStructure() {
  sheet_config.getRange('D2:D2').clearContent();

  let selectedAccount = sheet_config.getRange('C2:C2').getValue().split(" | ");
  let accountId = selectedAccount[0];
  let accountName = selectedAccount[1];

  if (accountId === "Select an account...") {
    sheet_config.getRange('D2:D2').setValue("<-- choose an account");
  } else {

    let proDropdown = [];

    let headers = ['accountId','accountName','propertyId','propertyName','propertyCreated','viewId','viewName'];
    data.push(headers);

    let responsePro = Analytics.Management.Webproperties.list(accountId);

    for (let i = 0; i < responsePro['totalResults']; i++) {
      let propertyId = responsePro['items'][i]['id'];
      let propertyName = responsePro['items'][i]['name'];
      let propertyCreated = responsePro['items'][i]['created'].split("T")[0];

      let responseVie = Analytics.Management.Profiles.list(accountId,propertyId);

      for (let j = 0; j < responseVie['totalResults']; j++) {
        let viewId = responseVie['items'][j]['id'];
        let viewName = responseVie['items'][j]['name'];
        
        data.push([accountId,accountName,propertyId,propertyName,propertyCreated,viewId,viewName]);
        proDropdown.push([`${propertyId} | ${propertyName}`]);
      }
    }

    writeData(sheet_accountStructure,data);
    sheet_hidden.getRange(3,3,1000,proDropdown[0].length).clearContent();
    sheet_hidden.getRange(3,3,proDropdown.length,proDropdown[0].length).setValues(proDropdown);

  }
}


// ---------- Output custom dimension details ----------
function getCustomDimensions() {
  sheet_config.getRange('D2:D2').clearContent();

  let selectedAccount = sheet_config.getRange('C2:C2').getValue().split(" | ");
  let accountId = selectedAccount[0];
  let accountName = selectedAccount[1];

  if (accountId === "Select an account...") {
    sheet_config.getRange('D2:D2').setValue("<-- choose an account");
  } else {

    let headers = ['accoundId','propertyId','propertyName','dimensionIndex','dimensionName','dimensionScope','dimensionActive','dimensionCreated'];
    data.push(headers);

    let responsePro = Analytics.Management.Webproperties.list(accountId);

    for (let i = 0; i < responsePro['totalResults']; i++) {
      let propertyId = responsePro['items'][i]['id'];
      let propertyName = responsePro['items'][i]['name'];

      let responseCus = Analytics.Management.CustomDimensions.list(accountId,propertyId);

      if (responseCus['totalResults'] === 0) {
        data.push([accountId,propertyId,propertyName,0,'','','','']);
      } else {
        for (let j = 0; j < responseCus['totalResults']; j++) {
          let dimensionIndex = responseCus['items'][j]['index'];
          let dimensionName = responseCus['items'][j]['name'];
          let dimensionScope = responseCus['items'][j]['scope'];
          let dimensionActive = responseCus['items'][j]['active'];
          let dimensionCreated = responseCus['items'][j]['created'].split("T")[0];

          data.push([accountId,propertyId,propertyName,dimensionIndex,dimensionName,dimensionScope,dimensionActive,dimensionCreated]);
        }
      }
    }

    writeData(sheet_customDimensions,data);
    sheet_customDimensionsMatrix.setTabColor('#0000FF');
    sheet_customDimensionsMatrix.autoResizeRows(1,1);
    sheet_customDimensionsMatrix.autoResizeColumns(1,1);
  }
}


// ---------- Output filter details ----------
function getFilters() {
  sheet_config.getRange('D2:D2').clearContent();

  let selectedAccount = sheet_config.getRange('C2:C2').getValue().split(" | ");
  let accountId = selectedAccount[0];
  let accountName = selectedAccount[1];

  if (accountId === "Select an account...") {
    sheet_config.getRange('D2:D2').setValue("<-- choose an account");
  } else {

    let headers = ['accountId','filterId','filterName','filterType','filterCreated','filterUpdated','filterField','filterMatchType','filterExpressionValue','filterCaseSensitive','filterFieldIndex','filterSearchString','filterReplaceString','filterFieldA','filterFieldAIndex','filterExtractA','filterFieldB','filterFieldBIndex','filterExtractB','filterOutputToField','filterOutputToFieldIndex','filterOutputConstructor','filterFieldARequired','filterFieldBRequired','filterOverrideOutputField'];

    data.push(headers);

    let responseFil = Analytics.Management.Filters.list(accountId);

    for (let i = 0; i < responseFil['totalResults']; i++) {
      let filterId = responseFil['items'][i]['id'];
      let filterName = responseFil['items'][i]['name'];
      let filterType = responseFil['items'][i]['type'];
      let filterCreated = responseFil['items'][i]['created'].split("T")[0];
      let filterUpdated = responseFil['items'][i]['updated'].split("T")[0];

      let detailTypeArr = filterType.toLowerCase().split("_");
      for (let k = 1; k < detailTypeArr.length; k++) {
        let x = detailTypeArr[k];
        detailTypeArr[k] = x.replace(x[0],x[0].toUpperCase());
      }
      let detailType = detailTypeArr.join('') + 'Details';

      let filterField = responseFil['items'][i][detailType]['field'] || '';
      let filterMatchType = responseFil['items'][i][detailType]['matchType'] || '';
      let filterExpressionValue = responseFil['items'][i][detailType]['expressionValue'] || '';
      let filterCaseSensitive = responseFil['items'][i][detailType]['caseSensitive'] || '';
      let filterFieldIndex = responseFil['items'][i][detailType]['fieldIndex'] || '';
      let filterSearchString = responseFil['items'][i][detailType]['searchString'] || '';
      let filterReplaceString = responseFil['items'][i][detailType]['replaceString'] || '';
      let filterFieldA = responseFil['items'][i][detailType]['fieldA'] || '';
      let filterFieldAIndex = responseFil['items'][i][detailType]['fieldAIndex'] || '';
      let filterExtractA = responseFil['items'][i][detailType]['extractA'] || '';
      let filterFieldB = responseFil['items'][i][detailType]['fieldB'] || '';
      let filterFieldBIndex = responseFil['items'][i][detailType]['fieldBIndex'] || '';
      let filterExtractB = responseFil['items'][i][detailType]['extractB'] || '';
      let filterOutputToField = responseFil['items'][i][detailType]['outputToField'] || '';
      let filterOutputToFieldIndex = responseFil['items'][i][detailType]['outputToFieldIndex'] || '';
      let filterOutputConstructor = responseFil['items'][i][detailType]['outputConstructor'] || '';
      let filterFieldARequired = responseFil['items'][i][detailType]['fieldARequired'] || '';
      let filterFieldBRequired = responseFil['items'][i][detailType]['fieldBRequired'] || '';
      let filterOverrideOutputField = responseFil['items'][i][detailType]['overrideOutputField'] || '';

      data.push([accountId,filterId,filterName,filterType,filterCreated,filterUpdated,filterField,filterMatchType,filterExpressionValue,filterCaseSensitive,filterFieldIndex,filterSearchString,filterReplaceString,filterFieldA,filterFieldAIndex,filterExtractA,filterFieldB,filterFieldBIndex,filterExtractB,filterOutputToField,filterOutputToFieldIndex,filterOutputConstructor,filterFieldARequired,filterFieldBRequired,filterOverrideOutputField]);
    }

    writeData(sheet_filters,data);

    getViewFilters();
    
  }
}


// ---------- Output filters linked to views ----------
function getViewFilters() {
  sheet_config.getRange('D2:D2').clearContent();

  let selectedAccount = sheet_config.getRange('C2:C2').getValue().split(" | ");
  let accountId = selectedAccount[0];
  let accountName = selectedAccount[1];

  data = [];

  if (accountId === "Select an account...") {
    sheet_config.getRange('D2:D2').setValue("<-- choose an account");
  } else {

    let headers = ['accountId','propertyId','viewId','viewName','filterId','filterName'];
    data.push(headers);

    if (sheet_config.getRange('C4:C4').getValue() !== 'all') {

      let propertyId = sheet_config.getRange('C4:C4').getValue().split(" | ")[0];

      let responseVie = Analytics.Management.Profiles.list(accountId,propertyId);

      for (let j = 0; j < responseVie['totalResults']; j++) {
        let viewId = responseVie['items'][j]['id'];
        let viewName = responseVie['items'][j]['name'];

        let responseVieFil = Analytics.Management.ProfileFilterLinks.list(accountId,propertyId,viewId);
        
        if (responseVieFil['totalResults'] === 0) {
          data.push([accountId,propertyId,viewId,viewName,0,'']);
        }

        for(k = 0; k < responseVieFil['totalResults']; k++) {
          let filterId = responseVieFil['items'][k]['filterRef']['id'];
          let filterName = responseVieFil['items'][k]['filterRef']['name'];

          data.push([accountId,propertyId,viewId,viewName,filterId,filterName]);
        }
      }

    } else {

      let responsePro = Analytics.Management.Webproperties.list(accountId);

      for (let i = 0; i < responsePro['totalResults']; i++) {
        let propertyId = responsePro['items'][i]['id'];

        let responseVie = Analytics.Management.Profiles.list(accountId,propertyId);

        for (let j = 0; j < responseVie['totalResults']; j++) {
          let viewId = responseVie['items'][j]['id'];
          let viewName = responseVie['items'][j]['name'];

          let responseVieFil = Analytics.Management.ProfileFilterLinks.list(accountId,propertyId,viewId);
          
          if (responseVieFil['totalResults'] === 0) {
            data.push([accountId,propertyId,viewId,viewName,0,'']);
          }

          for(k = 0; k < responseVieFil['totalResults']; k++) {
            let filterId = responseVieFil['items'][k]['filterRef']['id'];
            let filterName = responseVieFil['items'][k]['filterRef']['name'];

            data.push([accountId,propertyId,viewId,viewName,filterId,filterName]);
          }
        }
      }
    }

    console.log(data);
    writeData(sheet_viewFilters,data);
    buildViewFiltersMatrix(data);
  }
}


// ---------- Build view/filter matrix ----------
function buildViewFiltersMatrix(input) {
  sheet_viewFiltersMatrixPivot.getRange(1,1,sheet_viewFiltersMatrixPivot.getLastRow(),4).clear();

  let views = input;

  for (let i = 0; i < views.length; i++) {
    views[i].shift();
    let filterName = views[i].pop();
    let filterId = views[i].pop();
    if (filterId === 0) {
      views[i].push(filterId);
    } else {
      views[i].push(`${filterId} | ${filterName}`);
    }
  }
  //console.log(views);

  if (sheet_config.getRange('C4:C4').getValue() === 'all') {
    let filters = sheet_filters.getRange(2,2,sheet_filters.getLastRow()-1,2).getValues();
    let filterIdName = [];
    
    for (let i = 0; i < filters.length; i++) {
      filterIdName.push(`${filters[i][0]} | ${filters[i][1]}`);
    }

    let viewFilterIdName = [];

    for (let i = 1; i < views.length; i++) {
      if (!viewFilterIdName.includes(views[i][3])) {
        viewFilterIdName.push(views[i][3]);
      }
    }
    //console.log(viewFilterIdName);

    for (let i = 0; i < filterIdName.length; i++) {
      if (!viewFilterIdName.includes(filterIdName[i])) {
        views.push(['','','',filterIdName[i]]);
      }
    }


  }

  views.push(['','','',0]);
  console.log(views);
  writeData(sheet_viewFiltersMatrixPivot,views,"yes");
  
  sheet_viewFiltersMatrix.setTabColor('#0000FF');
  sheet_viewFiltersMatrix.autoResizeRows(1,1);
  sheet_viewFiltersMatrix.autoResizeColumns(1,1);
}


// ---------- Output goal details ----------
function getGoals() {
  sheet_config.getRange('D2:D2').clearContent();

  let selectedAccount = sheet_config.getRange('C2:C2').getValue().split(" | ");
  let accountId = selectedAccount[0];
  let accountName = selectedAccount[1];

  data = [];

  if (accountId === "Select an account...") {
    sheet_config.getRange('D2:D2').setValue("<-- choose an account");
  } else {

    let headers = ['accountId','propertyId','viewId','viewName','goalId','goalName','goalType','goalActive','goalCreated','goalUpdated','goalValue'];
    data.push(headers);

    if (sheet_config.getRange('C4:C4').getValue() === 'all') {

      let responsePro = Analytics.Management.Webproperties.list(accountId);

      for (let i = 0; i < responsePro['totalResults']; i++) {
        let propertyId = responsePro['items'][i]['id'];

        let responseVie = Analytics.Management.Profiles.list(accountId,propertyId);

        for (let j = 0; j < responseVie['totalResults']; j++) {
          let viewId = responseVie['items'][j]['id'];
          let viewName = responseVie['items'][j]['name'];

          let responseGoa = Analytics.Management.Goals.list(accountId,propertyId,viewId);

          for (let k = 0; k < responseGoa['totalResults']; k++) {
            let goalId = responseGoa['items'][k]['id'];
            let goalName = responseGoa['items'][k]['name'];
            let goalType = responseGoa['items'][k]['type'];
            let goalActive = responseGoa['items'][k]['active'];
            let goalCreated = responseGoa['items'][k]['created'].split("T")[0];
            let goalUpdated = responseGoa['items'][k]['updated'].split("T")[0];
            let goalValue = responseGoa['items'][k]['value'];

            data.push([accountId,propertyId,viewId,viewName,goalId,goalName,goalType,goalActive,goalCreated,goalUpdated,goalValue]);
          }
        }
      }

    } else {

      let propertyId = sheet_config.getRange('C4:C4').getValue().split(" | ")[0];

      let responseVie = Analytics.Management.Profiles.list(accountId,propertyId);

      for (let j = 0; j < responseVie['totalResults']; j++) {
        let viewId = responseVie['items'][j]['id'];
        let viewName = responseVie['items'][j]['name'];

        let responseGoa = Analytics.Management.Goals.list(accountId,propertyId,viewId);

        for (let k = 0; k < responseGoa['totalResults']; k++) {
          let goalId = responseGoa['items'][k]['id'];
          let goalName = responseGoa['items'][k]['name'];
          let goalType = responseGoa['items'][k]['type'];
          let goalActive = responseGoa['items'][k]['active'];
          let goalCreated = responseGoa['items'][k]['created'].split("T")[0];
          let goalUpdated = responseGoa['items'][k]['updated'].split("T")[0];
          let goalValue = responseGoa['items'][k]['value'];

          data.push([accountId,propertyId,viewId,viewName,goalId,goalName,goalType,goalActive,goalCreated,goalUpdated,goalValue]);
        }
      }
    }
    writeData(sheet_goals,data);
  }
}


function test() {
  
}