/**
 * Returns only the unique rows in the source array, discarding duplicates. 
 * The rows are returned according to the order in which they first appear in the source array.
 *
 * @param  {Object[][]} data a JavaScript 2d array
 * @param  {int} optColumnIndex the index of the column used to get unique values from a specific column
 * @param  {boolean} onlyReturnSelectedColumn a boolean, true to return only the values from optColumnIndex and no other column, false otherwise
 * @return {Object[][]} the unique rows in the source array, discarding duplicates. 
 */

function unique(data, optColumnIndex, onlyReturnSelectedColumn) {
  if (data.length > 0) {
    var o = {},
      i, l = data.length,
      r = [];
    if (typeof optColumnIndex == "number" && optColumnIndex < data[0].length) {
      if(onlyReturnSelectedColumn) {
        for (i = 0; i < l; i += 1) o[data[i][optColumnIndex]] = data[i][optColumnIndex];
      }
      else {
        for (i = 0; i < l; i += 1) o[data[i][optColumnIndex]] = data[i];
      }
      for (i in o) {
        if (o[i] != '') r.push(o[i]);
      }
    }
    else if (optColumnIndex == undefined) {
      for (i = 0; i < l; i += 1) o[data[i]] = data[i];
      for (i in o) {
        if (o[i] != '') r.push(o[i]);
      }
    }
    else {
      throw 'optColumnIndex should be a column index';
    }
    return r;
  }
  else {
    return data;
  }
}

/**
 * Returns the number of rows that meet certain criteria within a JavaScript 2d array.
 *
 * @param  {Object[][]} data a JavaScript 2d array
 * @param  {String} criteria the criteria in the form of a character string by which the cells are counted
 * @param  {boolean} matchEntireContent a boolean, true to match the entire cell content, false otherwise
 * @return {int} the number of rows that meet the criteria. 
 */

function countif(data, criteria, matchEntireContent) {
  if (data.length > 0) {
    var r = 0;
    var reg = new RegExp(escape(criteria));
    for (var i = 0; i < data.length; i++) {
      if(matchEntireContent){
        if (escape(data[i].toString()) == escape(criteria)) {
          r++;
        }
      }
      else{
        if (escape(data[i].toString()).search(reg) != -1) {
          r++;
        }
      }
    }
    return r;
  }
  else {
    return 0;
  }
}

/**
 * Returns a filtered version of the given source array, where only certain rows have been included.
 *
 * @param  {Object[][]} data a JavaScript 2d array
 * @param  {int} columnIndex the index of the column in which the values can be found or -1 to search accross all columns
 * @param  {String[]} values the values in the form of an array of strings 
 * @return {Object[][]} the filtered rows in the source array. 
 */ 

function filterByText(data, columnIndex, values) {
  var value = values;
  if (data.length > 0) {
    if (typeof columnIndex != "number" || columnIndex > data[0].length) {
      throw 'Choose a valide column index';
    }
    var r = [];
    if (typeof value == 'string') {
      var reg = new RegExp(escape(value).toUpperCase());
      for (var i = 0; i < data.length; i++) {
        if (columnIndex < 0 && escape(data[i].toString()).toUpperCase().search(reg) != -1 || columnIndex >= 0 && escape(data[i][columnIndex].toString()).toUpperCase().search(reg) != -1) r.push(data[i]);
      }
      return r;
    }
    else {
      for (var i = 0; i < data.length; i++) {
        for (var j = 0; j < value.length; j++) {
          var reg = new RegExp(escape(value[j]).toUpperCase());
          if (columnIndex < 0 && escape(data[i].toString()).toUpperCase().search(reg) != -1 || columnIndex >= 0 && escape(data[i][columnIndex].toString()).toUpperCase().search(reg) != -1) {
            r.push(data[i]);
            j = value.length;
          }
        }
      }
      return r;
    }
  }
  else {
    return data;
  }
}

/**
 * Returns the rows in the given array, sorted according to the given key column.
 *
 * @param  {Object[][]} data a JavaScript 2d array
 * @param  {int} columnIndex the index of the column to sort
 * @param  {boolean}  ascOrDesc a boolean, true for ascending, false for descending
 * @return {Object[][]} the sorted array. 
 */

function sort(data, columnIndex, ascOrDesc) {
  if (data.length > 0) {
    if (typeof columnIndex != "number" || columnIndex > data[0].length) {
      throw 'Choose a valide column index';
    }
    var r = new Array();
    var areDates = true;
    for (var i = 0; i < data.length; i++) {
      if(data[i] != null){ // 
        var value = data[i][columnIndex];
        if(value && typeof(value)=='string') { 
          var date = new Date(value);
          if (isNaN(date.getYear())) areDates = false;
          else data[i][columnIndex] = date;
        }
        r.push(data[i]);
      }
    }
    return r.sort(function (a, b) {
      if (ascOrDesc) return ((a[columnIndex] < b[columnIndex]) ? -1 : ((a[columnIndex] > b[columnIndex]) ? 1 : 0));
      return ((a[columnIndex] > b[columnIndex]) ? -1 : ((a[columnIndex] < b[columnIndex]) ? 1 : 0));
    });
  }
  else {
    return data;
  }
}