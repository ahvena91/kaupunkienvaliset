function main() {
  // get data
  var data = iterateRows();

  // get sorted 
  var henkkaritYleinen = getHenkkaritYleinen(data);
  var viisiottelutYleinen = getViisiottelutYleinen(data);
  var seitsenottelutYleinen = getSeitsenottelutYleinen(data);
  var joukkuekentatYleinen = getJoukkuekentatYleinen(data);
  var pystyhydratYleinen = getPystyhydratYleinen(data);
  var sviippitreenitYleinen = getSviippitreenitYleinen(data);

  var henkkaritNaiset = getHenkkaritNaiset(data);
  var viisiottelutNaiset = getViisiottelutNaiset(data);
  var seitsenottelutNaiset = getSeitsenottelutNaiset(data);
  var joukkuekentatNaiset = getJoukkuekentatNaiset(data);
  var pystyhydratNaiset = getPystyhydratNaiset(data);
  var sviippitreenitNaiset = getSviippitreenitNaiset(data);

  writeTulostaulu(henkkaritYleinen, henkkaritNaiset, "A","C");
  writeTulostaulu(viisiottelutYleinen, viisiottelutNaiset, "D","F");
  writeTulostaulu(seitsenottelutYleinen, seitsenottelutNaiset, "G","I");
  writeTulostaulu(joukkuekentatYleinen, joukkuekentatNaiset, "J","L");
  writeTulostaulu(pystyhydratYleinen, pystyhydratNaiset, "M","O");
  writeTulostaulu(sviippitreenitYleinen, sviippitreenitNaiset, "P","R");
}

// Writes sheet 'Tulostaulu' according to clumnStart and columnEnd positions
function writeTulostaulu(tulostauluYleinen, tulostauluNaiset, columnStart,columnEnd) {
  var rangeStringYleinen = "";
  var rangeStringNaiset = "";
  var sheet = SpreadsheetApp.getActive().getSheetByName("Tulostaulu");
  rangeStringYleinen = columnStart + "4:" + columnEnd + tulostauluYleinen.length.toString();
  rangeStringNaiset = columnStart + "35:" + columnEnd + tulostauluNaiset.length.toString();
  var clearRangeYleinen = sheet.getRange(columnStart+"4:"+columnEnd+"33");
  var clearRangeNaiset = sheet.getRange(columnStart+"35:"+columnEnd+"64");

  // clear old content
  clearRangeYleinen.clearContent();
  clearRangeNaiset.clearContent();
  const startRowYleinen = 4;
  const startRowNaiset = 35;

  // iterate length of result array (Yleinen)
  for (let i = 0; i<tulostauluYleinen.length; i++) {
    var j = i + startRowYleinen;
    rangeString = columnStart + j + ":" + columnEnd + j;
    // get range according to given range string
    var henkkaritRange = sheet.getRange(rangeString);
    var tmpArray = [[tulostauluYleinen[i].nimi, tulostauluYleinen[i].edustus, tulostauluYleinen[i].tulos]];
    // finally write values to given range
    henkkaritRange.setValues(tmpArray);
  }

  // iterate length of result array (Naiset)
  for (let i = 0; i<tulostauluNaiset.length; i++) {
    var j = i + startRowNaiset;
    rangeString = columnStart + j + ":" + columnEnd + j;
    // get range according to given range string
    var henkkaritRange = sheet.getRange(rangeString);
    var tmpArray = [[tulostauluNaiset[i].nimi, tulostauluNaiset[i].edustus, tulostauluNaiset[i].tulos]];
    // finally write values to given range
    henkkaritRange.setValues(tmpArray);
  }
}

// gets all data from "Pelatut tulokset" tab
function iterateRows() {
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

// returns array of henkkarit from series yleinen
function getHenkkaritYleinen(data) {
  var henkkaritYleinen = [];
  data.forEach(function (row) {
    if (row[2] == "Henkkari") {
      var tulosString = +row[5];
      var newHenkkari = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newHenkkari.sarja == "Yleinen sarja") {
        // iterate henkkaritYleinen for duplicates
        henkkaritYleinen.forEach(function (entry) {
          // found match in name/'nimi'
          if (newHenkkari.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newHenkkari.tulos > entry.tulos) {
              // use splice to cut old smaller result
              henkkaritYleinen.splice(henkkaritYleinen.indexOf(entry),1);
            }
          }
        });
        henkkaritYleinen.push(newHenkkari)
      }
    }
  });
  henkkaritYleinen.sort(compareDescOrder)
  return henkkaritYleinen;
}

// returns array of henkkarit from series naiset
function getHenkkaritNaiset(data) {
  var henkkaritNaiset = [];
  data.forEach(function (row) {
    if (row[2] == "Henkkari") {
      var tulosString = +row[5];
      var newHenkkari = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newHenkkari.sarja == "Naisten sarja") {
        // iterate henkkaritNaiset for duplicates
        henkkaritNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newHenkkari.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newHenkkari.tulos > entry.tulos) {
              // use splice to cut old smaller result
              henkkaritNaiset.splice(henkkaritNaiset.indexOf(entry),1);
            }
          }
        });
        henkkaritNaiset.push(newHenkkari)
      }
    }
  });
  henkkaritNaiset.sort(compareDescOrder)
  return henkkaritNaiset;
}

function getViisiottelutYleinen(data) {
  var viisiottelut = [];
  data.forEach(function (row) {
    if (row[2] == "5-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newViisiottelu = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newViisiottelu.sarja == "Yleinen sarja") {
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
    }
  });
  viisiottelut.sort(compareDescOrder)
  return viisiottelut;
}

function getViisiottelutNaiset(data) {
  var viisiottelutNaiset = [];
  data.forEach(function (row) {
    if (row[2] == "5-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newViisiottelu = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newViisiottelu.sarja == "Naisten sarja") {
        // iterate viisiottelutNaiset for duplicates
        viisiottelutNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newViisiottelu.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newViisiottelu.tulos > entry.tulos) {
              // use splice to cut old smaller result
              viisiottelutNaiset.splice(viisiottelutNaiset.indexOf(entry),1);
            }
          }
        });
        viisiottelutNaiset.push(newViisiottelu)
      }
    }
  });
  viisiottelutNaiset.sort(compareDescOrder)
  return viisiottelutNaiset;
}

function getSeitsenottelutYleinen(data) {
  var seitsenottelut = [];
  data.forEach(function (row) {
    if (row[2] == "7-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newSeitsenottelu = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newSeitsenottelu.sarja == "Yleinen sarja") {
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
    }
  });
  seitsenottelut.sort(compareDescOrder)
  return seitsenottelut;
}

function getSeitsenottelutNaiset(data) {
  var seitsenottelutNaiset = [];
  data.forEach(function (row) {
    if (row[2] == "7-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newSeitsenottelu = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newSeitsenottelu.sarja == "Naisten sarja") {
        // iterate seitsenottelutNaiset for duplicates
        seitsenottelutNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newSeitsenottelu.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newSeitsenottelu.tulos > entry.tulos) {
              // use splice to cut old smaller result
              seitsenottelutNaiset.splice(seitsenottelutNaiset.indexOf(entry),1);
            }
          }
        });
        seitsenottelutNaiset.push(newSeitsenottelu)
      }
    }
  });
  seitsenottelutNaiset.sort(compareDescOrder)
  return seitsenottelutNaiset;
}

function getJoukkuekentatYleinen(data) {
  var joukkuekentat = [];
  data.forEach(function (row) {
    if (row[2].includes("Joukkuekenttä")) {
      var tulosString = +row[5];
      var newJoukkuekentta = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newJoukkuekentta.sarja == "Yleinen sarja") {
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
    }
  });
  joukkuekentat.sort(compareDescOrder)
  return joukkuekentat;
}

function getJoukkuekentatNaiset(data) {
  var joukkuekentatNaiset = [];
  data.forEach(function (row) {
    if (row[2].includes("Joukkuekenttä")) {
      var tulosString = +row[5];
      var newJoukkuekentta = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newJoukkuekentta.sarja == "Naisten sarja") {
        // iterate joukkuekentatNaiset for duplicates
        joukkuekentatNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newJoukkuekentta.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newJoukkuekentta.tulos > entry.tulos) {
              // use splice to cut old smaller result
              joukkuekentatNaiset.splice(joukkuekentatNaiset.indexOf(entry),1);
            }
          }
        });
        joukkuekentatNaiset.push(newJoukkuekentta)
      }
    }
  });
  joukkuekentatNaiset.sort(compareDescOrder)
  return joukkuekentatNaiset;
}

function getPystyhydratYleinen(data) {
  var pystyhydrat = [];
  data.forEach(function (row) {
    if (row[2].includes("Pystyhydra")) {
      var tulosString = +row[5];
      var newPystyhydra = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newPystyhydra.sarja == "Yleinen sarja") {
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
    }
  });
  pystyhydrat.sort(compareAsceOrder)
  return pystyhydrat;
}

function getPystyhydratNaiset(data) {
  var pystyhydratNaiset = [];
  data.forEach(function (row) {
    if (row[2].includes("Pystyhydra")) {
      var tulosString = +row[5];
      var newPystyhydra = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newPystyhydra.sarja == "Naisten sarja") {
        // iterate pystyhydratNaiset for duplicates
        pystyhydratNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newPystyhydra.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newPystyhydra.tulos < entry.tulos) {
              // use splice to cut old smaller result
              pystyhydratNaiset.splice(pystyhydratNaiset.indexOf(entry),1);
            }
          }
        });
        pystyhydratNaiset.push(newPystyhydra)
      }
    }
  });
  pystyhydratNaiset.sort(compareAsceOrder)
  return pystyhydratNaiset;
}

function getSviippitreenitYleinen(data) {
  var sviippitreenit = [];
  data.forEach(function (row) {
    if (row[2].includes("Smuulin Sviippitreeni")) {
      var tulosString = +row[5];
      var newSviippitreeni = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newSviippitreeni.sarja == "Yleinen sarja") {
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
    }
  });
  sviippitreenit.sort(compareDescOrder)
  return sviippitreenit;
}

function getSviippitreenitNaiset(data) {
  var sviippitreenitNaiset = [];
  data.forEach(function (row) {
    if (row[2].includes("Smuulin Sviippitreeni")) {
      var tulosString = +row[5];
      var newSviippitreeni = {
        nimi : row[4],
        edustus : row[1],
        sarja : row[3],
        tulos : tulosString
      }
      // first check series
      if (newSviippitreeni.sarja == "Naisten sarja") {
        // iterate sviippitreenitNaiset for duplicates
        sviippitreenitNaiset.forEach(function (entry) {
          // found match in name/'nimi'
          if (newSviippitreeni.nimi == entry.nimi) {
            // if new result/'tulos' is greater than old one, remove old result
            if (newSviippitreeni.tulos > entry.tulos) {
              // use splice to cut old smaller result
              sviippitreenitNaiset.splice(sviippitreenitNaiset.indexOf(entry),1);
            }
          }
        });
        sviippitreenitNaiset.push(newSviippitreeni)
      }
    }
  });
  sviippitreenitNaiset.sort(compareDescOrder)
  return sviippitreenitNaiset;
}
