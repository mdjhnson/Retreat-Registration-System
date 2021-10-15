/*******************************
 TABLE OF CONTENT
   -  getReferences()
   -  getSubString(string, start, end)
   -  escapeRegExp(string)
   -  SStoNumColumn(name)
   -  
   

********************************/




/**
 * This stores reference information and passes it back to the function that called it.
 *              
 * @return A reference objects
 * 
 * @version 1.0 (09.01.21)
 * @author Matt Johnson <mdjhnson@gmail.com>
  */
function getReferences() {
  u = {
    settingHeaders: 1,

    //GENERAL INFO
    power_loc: {r: 11, c: 2},
    power: "",
    dataSourceURL_loc: {r: 2, c: 2},
    sheetName_loc: {r: 3, c: 2},
    dataHeaderRows_loc: {r: 4, c: 2},
    filterColumn_loc: {r: 5, c: 2},
    filterColumn: "",
    destinationURL_loc: {r: 6, c: 2},
    headerRows_loc: {r: 7,  c: 2},

    //FILTER LOGIC
    filterValue_loc: {r: 2, c: 6},
    filterSheetName_loc: {r: 2, c: 7},

    //DATA ROUTES
    sourceColumn_loc: {r: 2, c: 9},
    toColumn_loc: {r: 2, c: 10},
    findSubstring_loc: {r: 2, c: 11},
    startString_loc: {r: 2, c: 12},
    endString_loc: {r: 2, c:13}

  }
  u.power = settings[u.power_loc.r-1][u.power_loc.c-1]
  u.filterColumn = SStoNumColumn(settings[u.filterColumn_loc.r - 1][u.filterColumn_loc.c - 1]);


  return u;
}




/**
 * Match all characters between two strings
 *
 * @params {text} string - the string to search
 * @params {text} start - the string before desired text
 * @params {text} end - the string after desired text
 * @return {text} substring between values
 * 
 * @version 1.0 (10.14.21)
 * @author Matt Johnson <mdjhnson@gmail.com>
 */
function getSubString(string, start, end) {
  start = escapeRegExp(start)
  end = escapeRegExp(end)
  var patt = new RegExp("(?<=" + start + ")(.*)(?=" + end + ")", "gm")
  var result = string.match(patt)

  return result.join('; ')
}




function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}



/**
 * Converts letters to numbers. Often used to convert columns to array coordinates.
 * @param {string} name - The string of letters to be converted (ie. 'A' or 'AB').
 * @return numerical equivalent. 
 *              
 * @version 1.1 (03.09.21)
 * @author Matt JOhnson <mdjhnson@gmail.com>
 */
function SStoNumColumn(name) {
  //name = "AA"
  var i, index = 0;

  name = name.toUpperCase().split('');

  for (i = name.length - 1; i >= 0; i--) {
    var piece = name[i];
    var colNumber = piece.charCodeAt() - 64;
      index = index + colNumber * Number(Math.pow(26, name.length - (i + 1)));
  }
  return index || undefined;
}



function findLargest(arr) {
  var temp = 0;
  arr.forEach((element) => {
    if (temp < element) {
      temp = element;
    }
  });
  return temp;
}

