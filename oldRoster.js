function onOpen(){
    let ui = SpreadsheetApp.getUi();
  
    ui.createMenu('Dienstplan')
      .addItem('Mitarbeiterblätter erstellen', 'buildEmployeeSheets')
      .addItem('Dienstplan erstellen', 'buildRoster')
      .addToUi();
  }
  
  // ########################### Employees ###########################
  
  function buildEmployeeSheets(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const idSheetRange = ss.getSheetByName('Aktuelle Tabelle').getRange('A2');
  
    // checking if dates are correct
    const dateOne = ss.getSheetByName('Dienste Tage').getRange('A3').getValue();
    const dateTwo = ss.getSheetByName('Dienste Tage').getRange('B3').getValue();
  
    if(dateOne >= dateTwo){
      SpreadsheetApp.getUi().alert('Das zweite Datum muss älter als das erste Datum sein!');
    }else{
      // getting date for employeeSheet-Name
      const firstDate = dateOne.toLocaleString('de', { timeZone: 'Europe/Berlin', month: 'short', year: 'numeric' });
      
      const employeeSheetName = 'Dienstplan ' + firstDate;
      
      
      // creating new Spreadsheet
      const ssEmployee = SpreadsheetApp.create(employeeSheetName);
      const ssEmployeeId = ssEmployee.getId()
      idSheetRange.setValue(ssEmployeeId);
      ssEmployee.setSpreadsheetLocale('de');
      ssEmployee.setSpreadsheetTimeZone('Europe/Berlin');
  
      // creating holidaySheet
      const holidaySheet = ssEmployee.getActiveSheet().setName('Feiertage');
      const holidays = ss.getSheetByName('Feiertage').getDataRange().getValues();
      const holidayRangeDimension = ss.getSheetByName('Feiertage').getDataRange().getDataRegion().getA1Notation();
      holidaySheet.getRange(holidayRangeDimension).setValues(holidays);
      const dateRange = holidaySheet.getRange('B2:B20');
      ssEmployee.setNamedRange('Feiertage', dateRange);
      
      // creating EmployeeSheets
      const names = getNames();
      
      const firstNameSheet = ssEmployee.insertSheet().setName(names[0]);
      fillEmployeeSheet(ss, firstNameSheet);
  
      
      // duplicate first sheet
      names.slice(1).forEach(function(name){
        let newSheet = ssEmployee.duplicateActiveSheet();
        newSheet.setName(name);
        newSheet.getRange('B1:D2').setValue(name);
      }) 
  
      holidaySheet.hideSheet();
    }
  }
  
  function fillEmployeeSheet(spreadsheet, sheet){
    const datesheet = spreadsheet.getSheetByName('Dienste Tage');
    const name = sheet.getSheetName();
    
    insertingDates(sheet, datesheet, 2);
  
    // formating head
    const nameRange = sheet.getRange('B1:D2');
    nameRange.merge();
    nameRange.setValue(name);
    nameRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('silver').setFontSize(14).setFontWeight('bold');
    
    const serviceInfoRange = sheet.getRange('E1:K2');
    serviceInfoRange.merge();
    const serviceInfos = collectData(datesheet, 3, 4);
    let services = '';
    for(let i=0; i<serviceInfos.length; i++){
      if(i<serviceInfos.length-1){
        services += serviceInfos[i][1] + ': ' + serviceInfos[i][0] + ', ';
      }else{
        services += serviceInfos[i][1] + ': ' + serviceInfos[i][0];
      }
    }
    serviceInfoRange.setValue(services);
    serviceInfoRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('lightgrey');
    
    // formatting services
    serviceInfos.forEach(function(element,i){
      let servicesRange = sheet.getRange(4+i,1).setValue(element[1]);
      servicesRange.setBackground('silver').setFontWeight('bold');
    })
  
    // insert checkboxes
  
    const lc = sheet.getLastColumn();
    const checkboxRange = sheet.getRange(4,2,serviceInfos.length,lc-1);
    checkboxRange.insertCheckboxes();
    
    // deleting the checkboxes depending of weekdays or weekends
    let weekendDelete = [];
    serviceInfos.forEach(function(element, i){
      if(element[4] && !(element[5])){
        weekendDelete.push(i);
      }
    })
  
    let weekDayDelete = [];
    serviceInfos.forEach(function(element, i){
      if(element[5] && !(element[4])){
        weekDayDelete.push(i);
      }
    });
  
    for(let i=2;i<=lc;i++){
      
      if(sheet.getRange(3,i).getValue().getDay() > 4 || feiertage().includes(Date.parse(sheet.getRange(3,i).getValue()))){
        weekendDelete.forEach(element => sheet.getRange(element+4,i).removeCheckboxes());
      }else if(sheet.getRange(3,i).getValue().getDay() < 5 || !(feiertage().includes(Date.parse(sheet.getRange(3,i).getValue())))){
        weekDayDelete.forEach(element => sheet.getRange(element+4,i).removeCheckboxes());
      }
    }
  
    // last formatting
    sheet.setFrozenColumns(1);
    deletingEmptyRows(sheet);
  }
  
  
  function feiertage() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    const feiertageSheet = ss.getSheetByName('Feiertage');
    return feiertageSheet.getRange('B2:B13').getValues().map(date => Date.parse(date));
  }
  
  
  // ########################### Roster - Dienstplan ###########################
  
  function buildRoster(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeId = ss.getSheetByName('Aktuelle Tabelle').getRange('A2').getValue();
    const ssEmployee = SpreadsheetApp.openById(employeeId);
  
    creatingRoster(ssEmployee, getNames());
  }
  
  
  function creatingRoster(spreadsheet, names){ //Dienstplan erstellen mit allen Namen und Daten
    const sheet = spreadsheet.insertSheet('Dienstplan', 0);
    const dateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dienste Tage');
    
    // header Month: 
    // Datum aus dem Datumblatt auslesen und in den Monatskopf im Dienstplan eintragen
    const firstDate = dateSheet.getRange('A3');
    const monthRange = sheet.getRange('A1:B3').merge();
    monthRange.setValue(firstDate.getValue());
    
    monthRange.setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('silver').setFontSize(14).setFontWeight('bold');
    const dateFormat = "mmmm yyy";
    monthRange.setNumberFormat(dateFormat);
  
    // Names:
    // Die Namen auslesen und sie in die erste Spalte eintragen, immer mit einer Leerzeile dazwischen. Beginnend ab der vierten Zeile.
    names.forEach(function(name, i){
      sheet.getRange(4+(i*2),1).setValue(name);
    });
    
    sheet.autoResizeColumn(1);
  
    // Dates:
    // Die Daten vom Datenblatt eintragen und richtig formatieren
    insertingDates(sheet, dateSheet, 3);
    
    sheet.setFrozenColumns(2);
    deletingEmptyRows(sheet);
  
  
  
    const serviceInfos = collectData(dateSheet, 3, 4);
    Logger.log(serviceInfos);
    // Counter:
    // Counter für Anzahl Video/Audioschichten eintragen in das Feld neben dem Namen
    const counterRange = sheet.getRange('B4');
    const lastColumnA1Notation = sheet.getRange(5,sheet.getMaxColumns()).getA1Notation();
    let counterStringNotVideoTeam = '';
    for(let i=0;i<serviceInfos.length;i++){
      if(i === 0){
        counterStringNotVideoTeam += `="${serviceInfos[i][2].toUpperCase()}: "&COUNTIF(C5:${lastColumnA1Notation};"${serviceInfos[i][1]}") & `;
      }else{
        counterStringNotVideoTeam += `" ${serviceInfos[i][2].toUpperCase()}: "&COUNTIF(C5:${lastColumnA1Notation};"${serviceInfos[i][1]}") & `;
      }
      
    }
    for(let i=0;i<serviceInfos.length;i++){
      if(i === 0){
        counterStringNotVideoTeam += `" Stunden: " & COUNTIF(C5:${lastColumnA1Notation};"${serviceInfos[i][1]}")*${serviceInfos[i][3]} + ` ;
      }else if(i<serviceInfos.length-1){
        counterStringNotVideoTeam += `COUNTIF(C5:${lastColumnA1Notation};"${serviceInfos[i][1]}")*${serviceInfos[i][3]} + ` ;
      }else{
        counterStringNotVideoTeam += `COUNTIF(C5:${lastColumnA1Notation};"${serviceInfos[i][1]}")*${serviceInfos[i][3]}`;
      }
    }
    const counterString = `="V: "&COUNTIF(C5:${lastColumnA1Notation};"Tagschicht") + COUNTIF(C5:${lastColumnA1Notation};"Spät") + COUNTIF(C5:${lastColumnA1Notation};"Früh") & "(" & COUNTIF(C5:${lastColumnA1Notation};"Tagschicht") & ")" &" A: "&COUNTIF(C5:${lastColumnA1Notation};"Audio") & " Stunden: " & COUNTIF(C5:${lastColumnA1Notation};"Spät")*6 + COUNTIF(C5:${lastColumnA1Notation};"Früh")*6 + COUNTIF(C5:${lastColumnA1Notation};"Tagschicht")*8 + COUNTIF(C5:${lastColumnA1Notation};"Audio")*6`;
    counterRange.setValue(counterString);
    let destination = sheet.getRange(4,2,names.length*2,1);
    counterRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    for(let i=0;i<names.length;i++){
        let deleteRange = sheet.getRange(5+(i*2),2);
        deleteRange.clearContent();
    }
  
    // Formatting rows
    const lc = sheet.getLastColumn();
      
    for (let i=0; i<names.length; i++){
      if(i%2 == 0){
        let rowRange = sheet.getRange(4+(i*2), 1, 2, lc);
        rowRange.setBackground('lightgray');
      }else{
        let rowRange = sheet.getRange(4+(i*2), 1, 2, lc);
        rowRange.setBackground('aliceblue');
      }
      
      let chosenRange = sheet.getRange(5+(i*2), 3, 1, lc-2);
      chosenRange.setHorizontalAlignment('center');
      setChosenServiceFormatting(sheet, chosenRange, i);
    }
    
    
    // Fill in Data Validation, as source of Schichttypen
    const firstEmployeeSheet = spreadsheet.getSheetByName(names[0]);
    const schichtarten = firstEmployeeSheet.getRange(4,1,firstEmployeeSheet.getLastRow()-3);
    for(let i=0;i<names.length;i++){
      let destinationRange = sheet.getRange(5+(i*2),3,1, lc - 2);
      setServiceDataValidation(schichtarten, destinationRange)
    }
  
    sheet.autoResizeColumn(2);
  
  
    // Dataconnection
  
  
    for(let i=0;i<names.length;i++){
      let connectionString = '';
      for(let j=0;j<serviceInfos.length;j++){
        if(j===0){
          connectionString += `=IF(\'${names[i]}\'!B${j+4}; "${serviceInfos[j][2].toUpperCase()} "; "") `
        }else if(j===serviceInfos.length-1){
          connectionString += `& IF(\'${names[i]}\'!B${j+4}; "${serviceInfos[j][2].toUpperCase()}"; "")`
        }else{
          connectionString += `& IF(\'${names[i]}\'!B${j+4}; "${serviceInfos[j][2].toUpperCase()} "; "") `
        }
      }
      
      let nameRange = sheet.getRange(4+(i*2),3);
      nameRange.setValue(connectionString).setHorizontalAlignment('center');
      let destinationNames = sheet.getRange(4+(i*2), 3, 1, lc - 2);
      nameRange.autoFill(destinationNames, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
    function setChosenServiceFormatting(sheet, range, indexPerson){
      let upperRow = 4+(indexPerson*2);
      let downerRow = 5+(indexPerson*2);
    
      let warnRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=OR(AND(C${downerRow}="Spät"; NOT(REGEXMATCH(C${upperRow}; "S")));AND(C${downerRow}="Früh"; NOT(REGEXMATCH(C${upperRow}; "F")));AND(C${downerRow}="Tagschicht"; NOT(REGEXMATCH(C${upperRow}; "T")));AND(C${downerRow}="Audio"; NOT(REGEXMATCH(C${upperRow}; "A"))))`)
            .setBackground('Red')
            .setFontColor('white')
            .setRanges([range])
            .build();
  
      let spaetRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('Spät')
            .setBackground('Chocolate')
            .setFontColor('white')
            .setRanges([range])
            .build();
  
      let fruehRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('Früh')
            .setBackground('Coral')
            .setFontColor('white')
            .setRanges([range])
            .build();
  
      let tagRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('Tagschicht')
            .setBackground('Brown')
            .setFontColor('white')
            .setRanges([range])
            .build();
  
      let audioRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains('Audio')
            .setBackground('Blue')
            .setFontColor('white')
            .setRanges([range])
            .build();
  
      let newRules = [warnRule, spaetRule, fruehRule, tagRule, audioRule];
      setConditionalFormatting(sheet, newRules);
    }
    function setServiceDataValidation(sourceRange, destinationRange){ //Aufklappfelder
      let rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange).build();
      destinationRange.setDataValidation(rule);
    }
  }
  
  
  // ########################### Functions for Employees and Roster ###########################
  
  function setConditionalFormatting(sheet, newRules){ //alle Regeln die in newRules eingetragen sind, als bedingte Formatierung übergeben
    let rules = sheet.getConditionalFormatRules();
    newRules.forEach(element => rules.push(element));
    sheet.setConditionalFormatRules(rules);
  }
  
  
  function insertingDates(sheet, datesheet, startingColumn){ // Das Datum wird aus dem datesheet ausgelesen und im Zielblatt ab startingColumn eingetragen, formatiert und das conditional formatting intgriert
    const dateFrom = datesheet.getRange('A3').getValue();
    const dateTo = datesheet.getRange('B3').getValue();
    
    const sourceRange = sheet.getRange(3,startingColumn);
    sourceRange.setValue(dateFrom);
  
    // getting Total Days
    const totalDays = getTotalDays(dateFrom, dateTo);
    
    Logger.log(dateFrom);
    Logger.log(dateTo);
    Logger.log(totalDays);
  
    // formatting first Date
    insertColumns(sheet, totalDays, startingColumn);
    specialDateFormat(sourceRange);
    
  
    // autocomplete the rest of the dates
    // const destination = sheet.getRange(3, startingColumn, 1, totalDays+2);
    const destination = sheet.getRange(3, startingColumn, 1, totalDays);
    sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    // setting conditional formatting to the dates
    const range = sheet.getRange(3,startingColumn,1,sheet.getLastColumn()-1);
    setHolidayWeekendFormatting(sheet, range, startingColumn);
  
    function setHolidayWeekendFormatting(sheet, range, startingColumn){ //zwei bestimmte Formatierungen fürs Wochenende und für die Feiertage als bedingte Formatierung übergeben
      startingColumn = ['A','B','C','D','E','F','G'][startingColumn-1];
      let feiertageBedingteFormatierung = `=VLOOKUP(${startingColumn}3;INDIRECT("Feiertage");1;0)`;
      
      let holidayRule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(feiertageBedingteFormatierung)
          .setBackground("#FF0000")
          .setRanges([range])
          .build();
      
      let weekendRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains('S')
          .setBackground('orange')
          .setRanges([range])
          .build();
      
      let newRules = [holidayRule, weekendRule];
      setConditionalFormatting(sheet, newRules);
    }
    function specialDateFormat(range){
      let dateFormat = "ddd dd"; // "ddd dd" = "Mon 01"
      range.setNumberFormat(dateFormat);
  
      range.setHorizontalAlignment('center');
      range.setFontColor('white');
      range.setBackground('midnightblue');
    }
    function insertColumns(sheet, totalDays, startingColumn){
      let currentColumns = sheet.getMaxColumns();
      if(totalDays + startingColumn-1 > currentColumns){
        sheet.insertColumnsAfter(currentColumns, totalDays-currentColumns + (startingColumn-1));
      }
    }
    function getTotalDays(date1, date2){
      const diffTime = Math.abs(date2 - date1);
      return 1 + Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }
  }
  
  
  function deletingEmptyRows(sheet){
    sheet.deleteRows(sheet.getLastRow()+1, sheet.getMaxRows()-sheet.getLastRow());
  }
  
  
  function getNames(){
    const nameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mitarbeiter');
    return collectData(nameSheet,3,1,2).map(nameParts => nameParts[0] + ' ' + nameParts[1]);
  }
  
  function collectData(sheet, fromRow, fromColumn, sortedColumn){ //collects the data, intending to get all values of the sheet to right and down
    const range = sheet.getRange(fromRow, fromColumn, sheet.getLastRow() - fromRow +1, sheet.getLastColumn() - fromColumn +1);
    // Sorting the values before getting the values.
    if(sortedColumn != null){
      range.sort(sortedColumn);
    }
    return range.getValues();
  }