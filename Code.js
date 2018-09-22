// Import ES6 shim library
// ID: 1dvi84vwjD03YUc_-yp4D_fFjIXTK0J8Zk93qtNTPDI0xe2NO35XSv9em
// Learn more about this shim: http://ramblings.mcpher.com/Home/excelquirks/gassnips/es6shim
// Learn more about Libraries: https://developers.google.com/apps-script/guides/libraries
var Set = cEs6Shim.Set;
var Map = cEs6Shim.Map;


// Setup global vars
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetsCount = ss.getNumSheets();
var sheets = ss.getSheets();
var ui = SpreadsheetApp.getUi();

function onOpen() {
  
  ui.createMenu('RS Tools')
      .addItem('Parse Imported Data v1', 'parseTaxonomy')
      .addItem('Parse Imported Data v2', 'newParse')
      .addToUi();
}

/////////////////

function newParse() {
  
  //Clear cache
  SpreadsheetApp.flush;
  
  // Setup vars
  var parsedSheet = ss.getSheetByName("Parsed");
  var importSheet = ss.getSheetByName("Import");
  var numImportRows = importSheet.getLastRow();

  //Clear preexisting calues
  parsedSheet.getRange("A2:ZZ2000").clear(); 
  
  // Read Input
  var airtableImportArray = importSheet.getRange(1, 1, numImportRows, 1).getValues();
  // Parse Input calling parseTaxonomy, expect 2D array return in object
  var airtableParsedArray = (parseTaxonomyV2(airtableImportArray).airtableParsedArray);
  
  // Find nums of rows
  var airtableParsedArrayLenght = airtableParsedArray.length;
  // Find nums of cols in returned object
  var airtableParsedLongestLine = parseTaxonomyV2(airtableImportArray).longestLine;

  // Fill shorter rows with empty values
  airtableParsedArray.forEach(function(element) {
    while (element.length < airtableParsedLongestLine) {
      element.push("");
    }
  });
  
  // Establish target range
  var targetRange = parsedSheet.getRange(2,1,airtableParsedArrayLenght,airtableParsedLongestLine);

  // Write parsed data
  targetRange.setValues(Array.from(airtableParsedArray));

  //Clear cache
  SpreadsheetApp.flush;

}

function parseTaxonomyV2(airtableImportArray) {
  
  // Setup vars
  var airtableImportArrayLength = airtableImportArray.length;
  var airtableParsedArray = [[]];
  var airtableFlattenedArray = [];
  var longestLine = 0;

  // Iterate through lines calling parseLine, expect 2D array return
  for (i = 0; i < airtableImportArrayLength; i++){
    
    // Pass line to parseLine
    airtableParsedArray[i] = parseLine(airtableImportArray[i]);
    
    // Update longest Line record
    if (airtableParsedArray[i].length>longestLine){longestLine = airtableParsedArray[i].length;}
    
  }
  
  // Flatten 3D array
  airtableFlattenedArray = [].concat.apply([], airtableParsedArray);
  
  // Return onject with parsed data as 2D array and longest line integer
  return {
    airtableParsedArray: airtableFlattenedArray,
    longestLine: longestLine
  };
  
}

function parseLine(inputLineArray) {
  
  // Force input object into string (TODO better solution)
  var inputLineString = inputLineArray.toString();
  var inputLineLength = inputLineString.length;
  var outputLineArray = [[]];
  
  var cursorXY = [0,0];

  // Trenn-Symbole im String aus Airtable: |, ; und >
  var re = /\||\;|\>/;
  var regexSymbolTest = new RegExp(re);
    
  // Das Befüllen dieses Dictionaries mit Namen und Columns könnte man automatisieren aus dem Airtable Output
  var dictLabels = {
    "DIVISION": 0,
    "UNIT": 2,
    "GROUP": 8,
    "PROCESS": 13,
    "TASK": 17,
    "ROLE": 20,
    "ATTRIBUTE": 23,
  };
  
  // NOT IN USE
  var mappedArray = [[]];
  var splitArrayByHeader = [];
  











  // COPYPASTA

    
    


  ////// PARSING LOGIC //////
  /* |DIVISION>Corporate>Business Management#recH0zD5rv0NDTKo0;Business Development#rec4xjlDWZE1Xh4Tb;Operations Management#recuHdrREtJthLDix;Daily Business Operations#recZuLo60A4Z0tPZm;Technology#recsv9e76OBr1VRtN;Accounting#rec71Urt8q6EjRYqr;Legal#reczm08JzuO9c5nq8;Research and Development#recP0ecnpqsVhdOKZ;Purchasing#recBbYyMUJ9s9fkau;Workforce#recOkR5BHjlmgYHRP;Communications#recNQuA7pTjmipIQA;Customer Relationship Management#recP2qc9dkc5M6tG7|UNIT>Business Management#recH0zD5rv0NDTKo0>Vision#recMT43OMpI5zYIPs;Strategy#recYTGZzpgM0VFB1M;Stakeholder Management#recM4im4XgJ0EBg6m;Funding#recJOyn2MHGzOJLVc;Identity#rec3E4ZMHwRHPjm4y>Incident Handling#recp8OxRwDVAToT6L;Resource#recrhN5JGun1OrYYh>>Manager;Partner;Founder;Owner>Fee;Rental;Purchase|GROUP>Vision#recMT43OMpI5zYIPs>>Evaluation#reckFre47Qr3jtOQD;Implementation#recYEAGNCPp6XjzqI;Planning#recliczVnbZILYqzd;Development#recojzRxFYjzG6rXK;Research#recUfb5lF0MZhHMIN;Briefing#recpwmOpKGOiVqrc4>>|PROCESS>Sales Guidelines#recq6k9UiKNfFqAuv>Evaluation#reckFre47Qr3jtOQD;Implementation#recYEAGNCPp6XjzqI;Planning#recliczVnbZILYqzd;Development#recojzRxFYjzG6rXK;Research#recUfb5lF0MZhHMIN;Reporting#recxhjgK0ABIBaRIG;Briefing#recpwmOpKGOiVqrc4|TASK>Research#recUfb5lF0MZhHMIN>>|ROLE>Founder>>|ATTRIBUTE>Daily */
  
  for (i = 0 ; i <= airtableInputLineLen; i++){
    
    var thisChar = airtableInputString.charAt(i);
    var nextChar = airtableInputString.charAt(i+1);
    
    var rowOffset = 0;
    var previousRowOffset = 0;
    var fillCell;
    var printChar ="";
    var thisChar = "";
    var nextChar = "";
    var wordStart = 0;
    var wordEnd = 0;
    var thisWord = "";
    var charsUntilNextSymbol = 1;
    var skippedChars = 0;
    var columnOffset = 0;
  

    
    var thisChar = airtableInputLine.charAt(i);
    var nextChar = airtableInputLine.charAt(i+1);
    
    if(regexSymbolTest.test(thisChar) && regexSymbolTest.test(nextChar)) {
    
      skippedChars ++;
      if(thisChar === ">"){ columnOffset ++; }
      continue;
      
    }
    
    if(regexSymbolTest.test(thisChar) && !(regexSymbolTest.test(nextChar))) {
    
      if(thisChar === ">"){ columnOffset ++; }
      
      wordStart = i + 1;
      
      wordEnd = airtableInputLine.substr(wordStart, airtableInputLineLen - wordStart).search(re);
      
      thisWord = airtableInputLine.substr(wordStart, wordEnd);
      
      if(isInArray(Object.keys(dictLabels),thisWord) == false ) {
        
        rowOffset = i - skippedChars + previousRowOffset;

        //printCell.offset(rowOffset, columnOffset-1).setValue(thisWord);
        outputLineArray[rowOffset][columnOffset-1] = thisWord;

        if(rowOffset > 0 && columnOffset > 1) {
                    
          for(var key in dictLabels) {
            
            if(columnOffset-1<dictLabels[key]+1) {continue;}

            outputLineArray[rowOffset+2][dictLabels[key]+1] = outputLineArray[rowOffset+1][dictLabels[key]+1];
            
          }
                   
        }
        
      } else {
        
        columnOffset = dictLabels[thisWord];
        skippedChars ++;
      
      }
      
      skippedChars += thisWord.length;
      
    }
  }

   
  
  
  
  
  
  
  
  
  
  
  
  /* ATTEMPT at rewriting
  splitArrayByHeader = inputLineString.split("|");  

  mappedArray = splitArrayByHeader.map(function(x){x&"YADDA"});
  
  var mappedArray = splitArrayByHeader.map(function(level1) {
  
    level1 = level1.split(">");
    
    level1 = level1.map(function(level2) {
      level2 = level2.split(";");
      ui.alert(level2);
      return level2;
    });
    ui.alert(level1);
    return level1;
  });
  
  ui.alert(mappedArray[0]);
*/
  
/* RUDIMENT
splitArray1.forEach(function(element) {

    splitArray2.push(element.toString().split(">"));

  });*/
  
  /*// TEMP
  splitArrayByHeader = inputLineString.split("|");
  splitArrayByHeader.forEach(function(element){
    outputLineArray.push(element.split(">"));
  });*/
   
  //outputLineArray = mappedArray;
  
  // Return 2D array
  return outputLineArray;
}











function isInArray(array, search) {
  return array.indexOf(search) >= 0;
}














function parseTaxonomy() {
  
  var parsedSheet = ss.getSheetByName("Parsed");
  var importSheet = ss.getSheetByName("Import");
  
  // Trenn-Symbole im String aus Airtable: |, ; und >
  var re = /\||\;|\>/;
  var regexSymbolTest = new RegExp(re);
  
  // Das Befüllen dieses Dictionaries mit Namen und Columns könnte man automatisieren aus dem Airtable Output
  var dictLabels = {
    "DIVISION": 0,
    "UNIT": 2,
    "GROUP": 8,
    "PROCESS": 13,
    "TASK": 17,
    "ROLE": 20,
    "ATTRIBUTE": 23,
  };
  
  //Clear Values
  ss.getRange("A2:ZZ2000").clear(); 
  
  // Müssen diese deklarationen hier sein?
  var airtableInputRange = importSheet.getRange(1, 1, importSheet.getLastRow(), 1).getValues();
  var airtableInputRangeLen = airtableInputRange.length;
  
  var airtableInputLine;
  var airtableInputLineLen;

  var printCell = parsedSheet.getRange(2, 1);
  var rowOffset = 0;
  var previousRowOffset = 0;
  var fillCell;

  for (h = 1; h <= airtableInputRangeLen; h++) {
    
    var printChar ="";
    var thisChar = "";
    var nextChar = "";
    var wordStart = 0;
    var wordEnd = 0;
    var thisWord = "";
    var charsUntilNextSymbol = 1;
    var skippedChars = 0;
    var columnOffset = 0;
    
    airtableInputLine = importSheet.getRange(h, 1).getValue();
    airtableInputLineLen = airtableInputLine.length;

  for (i = 0 ; i <= airtableInputLineLen; i++){
    
    var thisChar = airtableInputLine.charAt(i);
    var nextChar = airtableInputLine.charAt(i+1);
    
    if(regexSymbolTest.test(thisChar) && regexSymbolTest.test(nextChar)) {
    
      skippedChars ++;
      if(thisChar === ">"){ columnOffset ++; }
      continue;
      
    }
    
    if(regexSymbolTest.test(thisChar) && !(regexSymbolTest.test(nextChar))) {
    
      if(thisChar === ">"){ columnOffset ++; }
      
      wordStart = i + 1;
      
      wordEnd = airtableInputLine.substr(wordStart, airtableInputLineLen - wordStart).search(re);
      
      thisWord = airtableInputLine.substr(wordStart, wordEnd);
      
      if(isInArray(Object.keys(dictLabels),thisWord) == false ) {
        
        rowOffset = i - skippedChars + previousRowOffset;

        printCell.offset(rowOffset, columnOffset-1).setValue(thisWord);
        
        if(rowOffset > 0 && columnOffset > 1) {
                    
          for(var key in dictLabels) {
            if(columnOffset-1<dictLabels[key]+1) {continue;}

            fillCell = parsedSheet.getRange(rowOffset+2, dictLabels[key]+1);
            fillCell.setValue(fillCell.offset(-1, 0).getValue());
          }
                   
        }
        
      } else {
        
        columnOffset = dictLabels[thisWord];
        skippedChars ++;
      
      }
      
      skippedChars += thisWord.length;
      
    }
  }
    previousRowOffset = rowOffset;
  }
}

function MD5 (input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  var txtHash = '';
  for (i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}