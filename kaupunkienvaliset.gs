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

// sort by result/'Tulos' in descending order
function compareDescOrder(a, b) {
  if(a.tulos > b.tulos)
    return -1;
  if(a.tulos < b.tulos)
    return 1;
  return 0;
}

// sort by result/'Tulos' in ascending oreder
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
      var newHenkkari = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate henkkarit for duplicates
      henkkarit.forEach(function (entry) {
        // found match in name/'nimi'
        if (newHenkkari.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newHenkkari.tulos > entry.tulos) {
            // use splice to cut old smaller result
            henkkarit.splice(henkkarit.indexOf(entry),1);
          }
        }
      });
      henkkarit.push(newHenkkari)
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
      var newViisiottelu = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate viisiottelut for duplicates
      viisiottelut.forEach(function (entry) {
        // found match in name/'nimi'
        if (newViisiottelu.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newViisiottelu.tulos > entry.tulos) {
            // use splice to cut old smaller result
            viisiottelut.splice(viisiottelut.indexOf(entry),1);
          }
        }
      });
      viisiottelut.push(newViisiottelu)
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
      var newSeitsenottelu = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate seitsenottelut for duplicates
      seitsenottelut.forEach(function (entry) {
        // found match in name/'nimi'
        if (newSeitsenottelu.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newSeitsenottelu.tulos > entry.tulos) {
            // use splice to cut old smaller result
            seitsenottelut.splice(seitsenottelut.indexOf(entry),1);
          }
        }
      });
      seitsenottelut.push(newSeitsenottelu)
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
      var newJoukkuekentta = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate joukkuekentat for duplicates
      joukkuekentat.forEach(function (entry) {
        // found match in name/'nimi'
        if (newJoukkuekentta.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newJoukkuekentta.tulos > entry.tulos) {
            // use splice to cut old smaller result
            joukkuekentat.splice(joukkuekentat.indexOf(entry),1);
          }
        }
      });
      joukkuekentat.push(newJoukkuekentta)
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
      var newPystyhydra = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate pystyhydrat for duplicates
      pystyhydrat.forEach(function (entry) {
        // found match in name/'nimi'
        if (newPystyhydra.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newPystyhydra.tulos < entry.tulos) {
            // use splice to cut old smaller result
            pystyhydrat.splice(pystyhydrat.indexOf(entry),1);
          }
        }
      });
      pystyhydrat.push(newPystyhydra)
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
      var newSviippitreeni = {
        nimi : row[4],
        edustus : row[1],
        tulos : tulosString
      }
      // iterate sviippitreenit for duplicates
      sviippitreenit.forEach(function (entry) {
        // found match in name/'nimi'
        if (newSviippitreeni.nimi == entry.nimi) {
          // if new result/'tulos' is greater than old one, remove old result
          if (newSviippitreeni.tulos > entry.tulos) {
            // use splice to cut old smaller result
            sviippitreenit.splice(sviippitreenit.indexOf(entry),1);
          }
        }
      });
      sviippitreenit.push(newSviippitreeni)
    }
  });
  sviippitreenit.sort(compareDescOrder)
  return sviippitreenit;
}
