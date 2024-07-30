function main() {
  // get data
  var data = iterateRows();
  // write results into 'Tulostaulu'
  writeTulostaulu(getHenkkaritYleinen(data), getHenkkaritNaiset(data), "A","C");
  writeTulostaulu(getViisiottelutYleinen(data), getViisiottelutNaiset(data), "D","F");
  writeTulostaulu(getSeitsenottelutYleinen(data), getSeitsenottelutNaiset(data), "G","I");
  writeTulostaulu(getJoukkuekentatYleinen(data), getJoukkuekentatNaiset(data), "J","L");
  writeTulostaulu(getPystyhydratYleinen(data), getPystyhydratNaiset(data), "M","O");
  writeTulostaulu(getSviippitreenitYleinen(data), getSviippitreenitNaiset(data), "P","R");

  // Write points for 'Henkkarit'
  writePistetaulu(getPisteet(getHenkkaritYleinen(data)), getPisteet(getHenkkaritNaiset(data)), "B");
  // Write points for '5o'
  writePistetaulu(getPisteet(getViisiottelutYleinen(data)), getPisteet(getViisiottelutNaiset(data)), "D");
  // Write points for '7o'
  writePistetaulu(getPisteet(getSeitsenottelutYleinen(data)), getPisteet(getSeitsenottelutNaiset(data)), "F");
  // Write points for 'JK'
  writePistetaulu(getPisteet(getJoukkuekentatYleinen(data)), getPisteet(getJoukkuekentatNaiset(data)), "H");
  // Write points for 'Pystyhydra'
  writePistetaulu(getPisteet(getPystyhydratYleinen(data)), getPisteet(getPystyhydratNaiset(data)), "J");
  // Write points for 'SmSv'
  writePistetaulu(getPisteet(getSviippitreenitYleinen(data)), getPisteet(getSviippitreenitNaiset(data)), "L");
}

// Writes sheet 'Pistetaulu' according to clumn positions
function writePistetaulu(pisteetY, pisteetN, column) {
  var rangeStringYleinen = "";
  var rangeStringNaiset = "";
  var sheet = SpreadsheetApp.getActive().getSheetByName("Pistetaulu");
  rangeStringYleinen = column + "8:" + column + "11";
  rangeStringNaiset = column + "14:" + column + "17";
  var clearRangeYleinen = sheet.getRange(rangeStringYleinen);
  var clearRangeNaiset = sheet.getRange(rangeStringNaiset);

  // clear old content
  clearRangeYleinen.clearContent();
  clearRangeNaiset.clearContent();
  const startRowYleinen = 8;
  const startRowNaiset = 14;

  // iterate length of points array (Yleinen)
  for (let i = 0; i<pisteetY.length; i++) {
    var j = i + startRowYleinen;
    rangeString = column + j;
    // get range according to given range string
    var henkkaritRange = sheet.getRange(rangeString);
    // finally write values to given range
    henkkaritRange.setValue(pisteetY[i]);
  }

  // iterate length of result array (Naiset)
  for (let i = 0; i<pisteetN.length; i++) {
    var j = i + startRowNaiset;
    rangeString = column + j;
    // get range according to given range string
    var henkkaritRange = sheet.getRange(rangeString);
    // finally write values to given range
    henkkaritRange.setValue(pisteetN[i]);
  }
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

// returns total points for a sport from series 'Yleinen'
function getPisteet(tulokset) {
  var pisteet = {
    pisteetTampere : 0,
    pisteetOulu : 0,
    pisteetLappee : 0,
    pisteetHelsinki : 0
  }

  var highestAmount = 30;
  var evenCasesArray = [];

  // iterate results 'Yleinen'
  for (let i=0; i<tulokset.length; i++) {
    if(i+1 < tulokset.length) {
      // is next result equal
      if(tulokset[i].tulos == tulokset[i+1].tulos) {
        var evenCase = {
          edustus : tulokset[i].edustus,
          pisteet : (highestAmount-i)
        }
        evenCasesArray.push(evenCase);

        if(i+2 < tulokset.length) {
          if(tulokset[i].tulos == tulokset[i+2].tulos) {
            continue;
          }
          else {
            var evenCase = {
              edustus : tulokset[i+1].edustus,
              pisteet : (highestAmount-i-1)
            }
            evenCasesArray.push(evenCase);
          }
        }
      }
      else {
        if(evenCasesArray == 0) {
          if(tulokset[i].edustus.includes("Länsi")) {
            pisteet.pisteetTampere = pisteet.pisteetTampere + (highestAmount-i);
          }
          else if(tulokset[i].edustus.includes("Pohjoinen")) {
            pisteet.pisteetOulu = pisteet.pisteetOulu + (highestAmount-i);
          }
          else if(tulokset[i].edustus.includes("Itä")) {
            pisteet.pisteetLappee = pisteet.pisteetLappee + (highestAmount-i);
          }
          else if(tulokset[i].edustus.includes("Etelä")) {
            pisteet.pisteetHelsinki = pisteet.pisteetHelsinki + (highestAmount-i);
          }
        }
        else {
          var temp = 0;
          evenCasesArray.forEach(function (evenCase) {
            temp = temp + evenCase.pisteet;
          });
          var meanValue = temp/evenCasesArray.length;
          evenCasesArray.forEach(function (evenCase) {
            if(evenCase.edustus.includes("Länsi")) {
              pisteet.pisteetTampere = pisteet.pisteetTampere + meanValue;
            }
            else if(evenCase.edustus.includes("Pohjoinen")) {
              pisteet.pisteetOulu = pisteet.pisteetOulu + meanValue;
            }
            else if(evenCase.edustus.includes("Itä")) {
              pisteet.pisteetLappee = pisteet.pisteetLappee + meanValue;
            }
            else if(evenCase.edustus.includes("Etelä")) {
              pisteet.pisteetHelsinki = pisteet.pisteetHelsinki + meanValue;
            }
          });
          evenCasesArray = [];
        }
      }
    }
    else {
      if(tulokset[i].edustus.includes("Länsi")) {
        pisteet.pisteetTampere = pisteet.pisteetTampere + (highestAmount-i);
      }
      else if(tulokset[i].edustus.includes("Pohjoinen")) {
        pisteet.pisteetOulu = pisteet.pisteetOulu + (highestAmount-i);
      }
      else if(tulokset[i].edustus.includes("Itä")) {
        pisteet.pisteetLappee = pisteet.pisteetLappee + (highestAmount-i);
      }
      else if(tulokset[i].edustus.includes("Etelä")) {
        pisteet.pisteetHelsinki = pisteet.pisteetHelsinki + (highestAmount-i);
      }
    }
  }

  // Arrange to array for Pistetaulu writing
  var pisteArray = [pisteet.pisteetOulu, pisteet.pisteetTampere, pisteet.pisteetLappee, pisteet.pisteetHelsinki];
  return pisteArray;
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
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "Henkkari") {
      var tulosString = +row[5];
      var newHenkkari = {
        nimi : row[4].trim(),
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
            else {
              duplicate = true;
            }
          }
          else {
            
          }
        });
        if (!duplicate) {
          henkkaritYleinen.push(newHenkkari)
          duplicate = false;
        }
      }
    }
  });
  henkkaritYleinen.sort(compareDescOrder)
  return henkkaritYleinen;
}

// returns array of henkkarit fromUntitled spreadsheet series naiset
function getHenkkaritNaiset(data) {
  var henkkaritNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "Henkkari") {
      var tulosString = +row[5];
      var newHenkkari = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          henkkaritNaiset.push(newHenkkari)
          duplicate = false;
        }
      }
    }
  });
  henkkaritNaiset.sort(compareDescOrder)
  return henkkaritNaiset;
}

function getViisiottelutYleinen(data) {
  var viisiottelut = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "5-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newViisiottelu = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          viisiottelut.push(newViisiottelu)
          duplicate = false;
        }
      }
    }
  });
  viisiottelut.sort(compareDescOrder)
  return viisiottelut;
}

function getViisiottelutNaiset(data) {
  var viisiottelutNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "5-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newViisiottelu = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          viisiottelutNaiset.push(newViisiottelu)
          duplicate = false;
        }
      }
    }
  });
  viisiottelutNaiset.sort(compareDescOrder)
  return viisiottelutNaiset;
}

function getSeitsenottelutYleinen(data) {
  var seitsenottelut = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "7-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newSeitsenottelu = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          seitsenottelut.push(newSeitsenottelu)
          duplicate = false;
        }
      }
    }
  });
  seitsenottelut.sort(compareDescOrder)
  return seitsenottelut;
}

function getSeitsenottelutNaiset(data) {
  var seitsenottelutNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2] == "7-ottelu") {
      var tulosString = row[5].toString().replace(",",".");
      tulosString = +tulosString;
      var newSeitsenottelu = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          seitsenottelutNaiset.push(newSeitsenottelu)
          duplicate = false;
        }
      }
    }
  });
  seitsenottelutNaiset.sort(compareDescOrder)
  return seitsenottelutNaiset;
}

function getJoukkuekentatYleinen(data) {
  var joukkuekentat = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Joukkuekenttä")) {
      var tulosString = +row[5];
      var newJoukkuekentta = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          joukkuekentat.push(newJoukkuekentta)
          duplicate = false;
        }
      }
    }
  });
  joukkuekentat.sort(compareDescOrder)
  return joukkuekentat;
}

function getJoukkuekentatNaiset(data) {
  var joukkuekentatNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Joukkuekenttä")) {
      var tulosString = +row[5];
      var newJoukkuekentta = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          joukkuekentatNaiset.push(newJoukkuekentta)
          duplicate = false;
        }
      }
    }
  });
  joukkuekentatNaiset.sort(compareDescOrder)
  return joukkuekentatNaiset;
}

function getPystyhydratYleinen(data) {
  var pystyhydrat = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Pystyhydra")) {
      var tulosString = +row[5];
      var newPystyhydra = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          pystyhydrat.push(newPystyhydra)
          duplicate = false;
        }
      }
    }
  });
  pystyhydrat.sort(compareAsceOrder)
  return pystyhydrat;
}

function getPystyhydratNaiset(data) {
  var pystyhydratNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Pystyhydra")) {
      var tulosString = +row[5];
      var newPystyhydra = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          pystyhydratNaiset.push(newPystyhydra)
          duplicate = false;
        }
      }
    }
  });
  pystyhydratNaiset.sort(compareAsceOrder)
  return pystyhydratNaiset;
}

function getSviippitreenitYleinen(data) {
  var sviippitreenit = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Smuulin Sviippitreeni")) {
      var tulosString = +row[5];
      var newSviippitreeni = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          sviippitreenit.push(newSviippitreeni)
          duplicate = false;
        }
      }
    }
  });
  sviippitreenit.sort(compareDescOrder)
  return sviippitreenit;
}

function getSviippitreenitNaiset(data) {
  var sviippitreenitNaiset = [];
  var duplicate = false;
  data.forEach(function (row) {
    if (row[2].includes("Smuulin Sviippitreeni")) {
      var tulosString = +row[5];
      var newSviippitreeni = {
        nimi : row[4].trim(),
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
        if (!duplicate) {
          sviippitreenitNaiset.push(newSviippitreeni)
          duplicate = false;
        }
      }
    }
  });
  sviippitreenitNaiset.sort(compareDescOrder)
  return sviippitreenitNaiset;
}
