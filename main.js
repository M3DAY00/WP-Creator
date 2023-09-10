function doGet() {

  var files = DriveApp.getFilesByName("Paramétrage (M3DAY00-AS)");
  
  while (files.hasNext()) {
    var file = files.next();
    var sheetId = file.getId();
    var sSheet = SpreadsheetApp.openById(sheetId);
    var sheet1 = sSheet.getSheets()[0];
    var sheet2 = sSheet.getSheets()[1];
    var sheet3 = sSheet.getSheets()[2];
  }
  if (!sheetId) {
    var createdSheet = SpreadsheetApp.create("Paramétrage (M3DAY00-AS)");
    var sheetId = createdSheet.getId();
    var sSheet = SpreadsheetApp.openById(sheetId);
    var sheet1 = sSheet.getSheets()[0];
    sheet1.getRange("1:1000").setVerticalAlignment("middle");
    sheet1.deleteColumns(3, 24);
    sheet1.deleteRows(4, 997);
    sheet1.setName("Début/fin");
    sheet1.getRange("A1:B1").merge();
    sheet1.setColumnWidth(1, 200);
    sheet1.setColumnWidth(2, 200);
    sheet1.getRange("A1").setValue("Début/fin de l'année scolaire (AAAA-MM-JJ)");
    sheet1.getRange("A1").setHorizontalAlignment("center");
    sheet1.getRange("A2").setValue("Début :");
    sheet1.getRange("A2").setHorizontalAlignment("right");
    sheet1.getRange("A3").setValue("Fin :");
    sheet1.getRange("A3").setHorizontalAlignment("right");
    var sheet2 = sSheet.insertSheet();
    sheet2.getRange("1:1000").setVerticalAlignment("middle");
    sheet2.deleteColumns(2, 25);
    sheet2.setColumnWidth(1, 400);
    sheet2.setName("Congés/pédagogiques");
    sheet2.getRange("A1").setValue("Congés/pédagogiques (AAAA-MM-JJ)");
    sheet2.getRange("A1").setHorizontalAlignment("center");
    var sheet3 = sSheet.insertSheet();
    sheet3.getRange("1:1000").setVerticalAlignment("middle");
    sheet3.deleteColumns(2, 25);
    sheet3.setColumnWidth(1, 100);
    sheet3.setName("Groupes");
    sheet3.getRange("A1").setValue("Groupes");
    sheet3.getRange("A:A").setHorizontalAlignment("center");
    return "Un Google Spreadsheet à été créé. Veuillez remplir les données pour créer les documents.\nLe Spreadsheet est disponible à l'addresse suivante : " + createdSheet.getUrl()
  } else {
    if (sSheet.getSheets().length == 3) {
      if (sheet1.getRange("B2").getValue() != "" && sheet1.getRange("B3").getValue() != "" && sheet2.getRange("A2").getValue() != "" && sheet3.getRange("A2").getValue() != "") {
        sSheet.getSheets()[2].activate();
        var sheet4 = sSheet.insertSheet();
        sheet4.deleteColumns(2, 25);
        sheet4.deleteRows(19, 981);
        var groups = sheet3.getRange("A2:A").getValues();
        for (var i = 0; i < parseInt(groups.length); i++) {
          if (groups[i][0] !== "") {
            sheet4.insertColumns(i + 1);
            sheet4.setColumnWidth(i + 1, 100);
            sheet4.getRange(1, i + 1).setValue(sheet3.getRange(i + 2, 1).getValue());
          }
        }
        sheet4.insertColumns(1);
        sheet4.setName("Jours");
        sheet4.setColumnWidth(1, 100);
        sheet4.getRange("A2").setValue("Jour 1 :");
        sheet4.getRange("A2").setHorizontalAlignment("right");
        sheet4.getRange("A3").setValue("Jour 2 :");
        sheet4.getRange("A3").setHorizontalAlignment("right");
        sheet4.getRange("A4").setValue("Jour 3 :");
        sheet4.getRange("A4").setHorizontalAlignment("right");
        sheet4.getRange("A5").setValue("Jour 4 :");
        sheet4.getRange("A5").setHorizontalAlignment("right");
        sheet4.getRange("A6").setValue("Jour 5 :");
        sheet4.getRange("A6").setHorizontalAlignment("right");
        sheet4.getRange("A7").setValue("Jour 6 :");
        sheet4.getRange("A7").setHorizontalAlignment("right");
        sheet4.getRange("A8").setValue("Jour 7 :");
        sheet4.getRange("A8").setHorizontalAlignment("right");
        sheet4.getRange("A9").setValue("Jour 8 :");
        sheet4.getRange("A9").setHorizontalAlignment("right");
        sheet4.getRange("A10").setValue("Jour 9 :");
        sheet4.getRange("A10").setHorizontalAlignment("right");
        sheet4.getRange("A11").setValue("Jour 10 :");
        sheet4.getRange("A11").setHorizontalAlignment("right");
        sheet4.getRange("A12").setValue("Jour 11 :");
        sheet4.getRange("A12").setHorizontalAlignment("right");
        sheet4.getRange("A13").setValue("Jour 12 :");
        sheet4.getRange("A13").setHorizontalAlignment("right");
        sheet4.getRange("A14").setValue("Jour 13 :");
        sheet4.getRange("A14").setHorizontalAlignment("right");
        sheet4.getRange("A15").setValue("Jour 14 :");
        sheet4.getRange("A15").setHorizontalAlignment("right");
        sheet4.getRange("A16").setValue("Jour 15 :");
        sheet4.getRange("A16").setHorizontalAlignment("right");
        sheet4.getRange("A17").setValue("Jour 16 :");
        sheet4.getRange("A17").setHorizontalAlignment("right");
        sheet4.getRange("A18").setValue("Jour 17 :");
        sheet4.getRange("A18").setHorizontalAlignment("right");
        sheet4.getRange("A19").setValue("Jour 18 :");
        sheet4.getRange("A19").setHorizontalAlignment("right");
        sheet4.deleteColumn(sheet4.getMaxColumns())
      }
      else {
        return "Le script n'a pas été en mesure de créer une plage pour assigner les jours aux groupes puisque vous n\'avez pas remplis les cases nécessaires au fonctionnement."
      }
    }
    else if (sSheet.getSheets().length == 4) {
      CreateDoc()
    }
    else {
      return "La structure du Google Spreadsheet à été modifiés et ne peut donc pas être pris en charge. Veuillez sauvegarder vos données, supprimer le Google Spreadsheet et de revenir sur rafraichir cette page."
    }
  }
}

function CreateDoc() {
  var files = DriveApp.getFilesByName("Paramétrage (M3DAY00-AS)");
  while (files.hasNext()) {
    var file = files.next();
    var sheetId = file.getId();
    var sSheet = SpreadsheetApp.openById(sheetId);
    var sheets = sSheet.getSheets();
  }
  var sheetId = sSheet.getId();
  var sSheet = SpreadsheetApp.openById(sheetId);
  var groupsInt = sheets[3].getMaxColumns() - 1;
  var groups = {}
  for (var a = 1; a <= groupsInt; a++) {
    var groupNo = sheets[3].getRange(1, a + 1).getValue()
    groups[groupNo] = []
    for (j = 2; j <= 19; j++) {
      groups[groupNo][j - 2] = sheets[3].getRange(j, a + 1).getValue()
    }
    var doc0 = DocumentApp.create(groupNo + " PT");
    var body0 = doc0.getBody();
    body0.appendTable([["", "Pendant le cours", "Devoir(s)"]]);
    var day = 18;
    var daysFromStart = 0;
    for (var i = new Date(sheets[0].getRange("B2").getValue()); i <= new Date(sheets[0].getRange("B3").getValue()); i.setDate(i.getDate() + 1))
    {
      if(i.getDay() > 0 && i.getDay() < 6) {
        isDayOff = false
        for (var x = 2; x <= sheets[1].getMaxRows(); x++) {
          if (sheets[1].getRange(x, 1).getValue() == "") {
            break;
          }
          dateCheck = new Date(sheets[1].getRange(x, 1).getValue());
          if (dateCheck == i) {
            isDayOff = true;
            break;
          }
        }
        if (isDayOff == false) {
          if (day > 17) {
            day = 0
          } else {
            day++;
          }
          if (groups[groupNo][day] != "") {
            daysFromStart++;
            var table = body0.getTables()[0];
            var newRow = table.appendTableRow();
            newRow.appendTableCell("Cours " + daysFromStart + "\n\n" + FormatD(i));
            newRow.appendTableCell('');
            newRow.appendTableCell('');
          }
        }
      }
    }
    doc0.saveAndClose();
  }
  return "Les documents ont été complétés. Visiter Google Docs pour les voirs et les modifiés!"
}

function FormatD(dateToFormat) {
  const date = new Date(dateToFormat);
  let strMonth = "";
  if(date.getMonth() == 0)
  {
    strMonth = "janvier";
  }
  else if(date.getMonth() == 1)
  {
    strMonth = "février";
  }
  else if(date.getMonth() == 2)
  {
    strMonth = "mars";
  }
  else if(date.getMonth() == 3)
  {
    strMonth = "avril";
  }
  else if(date.getMonth() == 4)
  {
    strMonth = "mai"
  }
  else if(date.getMonth() == 5)
  {
    strMonth = "juin"
  }
  else if(date.getMonth() == 6)
  {
    strMonth = "juillet"
  }
  else if(date.getMonth() == 7)
  {
    strMonth = "août"
  }
  else if(date.getMonth() == 8)
  {
    strMonth = "septembre"
  }
  else if(date.getMonth() == 9)
  {
    strMonth = "octobre"
  }
  else if(date.getMonth() == 10)
  {
    strMonth = "novembre"
  }
  else if(date.getMonth() == 11)
  {
    strMonth = "décembre"
  }
  if (date.getDay() == 0) {
    strDay = "Dimanche";
  } else if (date.getDay() == 1) {
    strDay = "Lundi";
  } else if (date.getDay() == 2) {
    strDay = "Mardi";
  } else if (date.getDay() == 3) {
    strDay = "Mercredi";
  } else if (date.getDay() == 4) {
    strDay = "Jeudi";
  } else if (date.getDay() == 5) {
    strDay = "Vendredi";
  } else if (date.getDay() == 6) {
    strDay = "Samedi";
  } 
  dayOfMonth = date.getDate()
  return strDay + "\n" + dayOfMonth + " " + strMonth
}
