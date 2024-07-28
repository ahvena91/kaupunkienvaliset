function main() {
  var data = iterateThroughRows();
  var henkkarit = getHenkkari(data);
  var viisiottelut = getViisiottelu(data);
  var seitsenottelut = getSeitsenottelu(data);
  var joukkuekentat = getJoukkuekentta(data);
  var pystyhydrat = getPystyhydra(data);
  var sviippitreenit = getSviippitreeni(data);
  writeTulostaulu(henkkarit,"A","C");
  writeTulostaulu(viisiottelut,"D","F");
  writeTulostaulu(seitsenottelut,"G","I");
  writeTulostaulu(joukkuekentat,"J","L");
  writeTulostaulu(pystyhydrat,"M","O");
  writeTulostaulu(sviippitreenit,"P","R");
}

// Writes sheet 'Tulostaulu' according to clumnStart and columnEnd positions
function writeTulostaulu(tulostaulu,columnStart,columnEnd) {
  var rangeString = "";
  var sheet = SpreadsheetApp.getActive().getSheetByName("Tulostaulu");
  rangeString = columnStart + "3:" + columnEnd + tulostaulu.length.toString();
  var clearRange = sheet.getRange(columnStart+"3:"+columnEnd+"32");
  clearRange.clearContent();
  const startRow = 3;
  // iterate lenght of result array
  for (let i = 0; i<tulostaulu.length; i++) {
    var j = i + startRow;
    rangeString = columnStart + j + ":" + columnEnd + j;
    // get range according to given range string
    var henkkaritRange = sheet.getRange(rangeString);
    var tmpArray = [[tulostaulu[i].nimi, tulostaulu[i].edustus, tulostaulu[i].tulos]];
    // finally write values to given range
    henkkaritRange.setValues(tmpArray);
  }
}

// gets all data from "Pelatut tulokset" tab
function iterateThroughRows() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Pelatut tulokset");
  var data = sheet.getDataRange().getValues();
  return data;
}

// sort by 'Tulos'
function compareDescOrder(a, b) {
  if(a.tulos > b.tulos)
    return -1;
  if(a.tulos < b.tulos)
    return 1;
  return 0;
}

// sort by 'Tulos'
function compareAsceOrder(a, b) {
  if(a.tulos < b.tulos)
    return -1;
  if(a.tulos > b.tulos)
    return 1;
  return 0;
}

function getHenkkari(data) {
  var henkkarit = [];
  data.forEach(function (row) {
    if (row[2] == "Henkkari") {
      var tulosString = +row[5];
      var henkkari = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      henkkarit.push(henkkari)
    }
  });
  henkkarit.sort(compareDescOrder)
  return henkkarit;
}

function getViisiottelu(data) {
  var viisiottelut = [];
  data.forEach(function (row) {
    if (row[2] == "5-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var viisiottelu = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      viisiottelut.push(viisiottelu)
    }
  });
  viisiottelut.sort(compareDescOrder)
  return viisiottelut;
}

function getSeitsenottelu(data) {
  var seitsenottelut = [];
  data.forEach(function (row) {
    if (row[2] == "7-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var seitsenottelu = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      seitsenottelut.push(seitsenottelu)
    }
  });
  seitsenottelut.sort(compareDescOrder)
  return seitsenottelut;
}

function getJoukkuekentta(data) {
  var joukkuekentat = [];
  data.forEach(function (row) {
    if (row[2].includes("Joukkuekenttä")) {
      var tulosString = +row[5];
      var joukkuekentta = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      joukkuekentat.push(joukkuekentta)
    }
  });
  joukkuekentat.sort(compareDescOrder)
  return joukkuekentat;
}

function getPystyhydra(data) {
  var pystyhydrat = [];
  data.forEach(function (row) {
    if (row[2].includes("Pystyhydra")) {
      var tulosString = +row[5];
      var pystyhydra = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      pystyhydrat.push(pystyhydra)
    }
  });
  pystyhydrat.sort(compareAsceOrder)
  return pystyhydrat;
}

function getSviippitreeni(data) {
  var sviippitreenit = [];
  data.forEach(function (row) {
    if (row[2].includes("Smuulin Sviippitreeni")) {
      var tulosString = +row[5];
      var sviippitreeni = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      sviippitreenit.push(sviippitreeni)
    }
  });
  sviippitreenit.sort(compareDescOrder)
  return sviippitreenit;
}
