//By hirobank, last edited on 19-Jan-2021

class easySheet {
  constructor(sheet) { //Parameter must be a *Sheet* object but NOT a string
      const range = sheet.getDataRange();
      const values = range.getValues();
      this.idSeriesArray = values[0].slice(1);
      for (var line = 1; line < values.length ; line++) {
        var key = { 'active' : 1, }
        for (var col = 1; col < values[0].length ; col++ )
            key[values[0][col]] = values[line][col];
        this[values[line][0]] = key;
      }
      var errorMsg = this.CheckIdentifierUnique();
    if ( errorMsg != 0 ) {
      Browser.msgBox(errorMsg);
    }
  }

  getNumByRow(sheet,idSeries,idRow) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const keysArray = Object.keys(this);
    var numRow = { 'errorCode' : 0, 'keysMissing' : [], 'keysDuplicated' : [], };
    //Error Code: Normal - 0, Missing Only - 1, Duplicated Only - 2, Both - 3

    for (var countKey = 0; countKey < keysArray.length; countKey++) {
      var countFound = 0;
      if (keysArray[countKey] == 'idSeriesArray')
        continue;
      for (var row = 0; row < values[0].length; row++) {
        if (this[keysArray[countKey]][idSeries] == values[idRow][row] && this[keysArray[countKey]][idSeries] !='__skip__') {
          numRow[keysArray[countKey]] = row;
          countFound += 1;
        }
      }
      if (countFound == 0  && this[keysArray[countKey]][idSeries] !='__skip__')
        numRow.keysMissing.push(keysArray[countKey]);
      if (countFound > 1  && this[keysArray[countKey]][idSeries] !='__skip__')
        numRow.keysDuplicated.push(keysArray[countKey]);
    }
    if (numRow.keysMissing.length != 0)
      numRow.errorCode += 1;
    if (numRow.keysDuplicated.length != 0)
      numRow.errorCode += 2;
    return numRow;
  }

  getNumByCol(sheet,idSeries,idCol) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const keysArray = Object.keys(this);
    var numCol = { 'errorCode' : 0, 'keysMissing' : [], 'keysDuplicated' : [], };
    //Error Code: Normal - 0, Missing Only - 1, Duplicated Only - 2, Both - 3

    for (var countKey = 0; countKey < keysArray.length; countKey++) {
      var countFound = 0;
      if (keysArray[countKey] == 'idSeriesArray')
        continue;
      for (var col = 0; col < values.length; col++) {
        if (this[keysArray[countKey]][idSeries] == values[col][idCol] && this[keysArray[countKey]][idSeries] !='__skip__') {
          numCol[keysArray[countKey]] = col;
          countFound += 1;
        }
      }
      if (countFound == 0  && this[keysArray[countKey]][idSeries] !='__skip__')
        numCol.keysMissing.push(keysArray[countKey]);
      if (countFound > 1  && this[keysArray[countKey]][idSeries] !='__skip__')
        numCol.keysDuplicated.push(keysArray[countKey]);
    }
    if (numCol.keysMissing.length != 0)
      numCol.errorCode += 1;
    if (numCol.keysDuplicated.length != 0)
      numCol.errorCode += 2;
    return numCol;
  }

  CheckIdentifierUnique() {
    var errorCode = 0;
    var errorMsg = '';
    
    var keysResult = this.duplicateCheck(Object.keys(this));
    errorCode += keysResult.errorCode;
    errorMsg += keysResult.errorMsg;
    var seriesResult = this.duplicateCheck(Object.keys(this.idSeriesArray));
    errorCode += seriesResult.errorCode;
    errorMsg += seriesResult.errorMsg;

    for (var countId = 0; countId < this.idSeriesArray.length; countId++) {
      var checkArray = [];
      var idSeries = this.idSeriesArray[countId]
      var keysArray = Object.keys(this);
      for (var countKey = 0; countKey < keysArray.length; countKey++) {
        var keys = keysArray[countKey];
        if (keys != 'idSeries'  && this[keys][idSeries] !='__skip__')
          checkArray.push(this[keys][idSeries]);
      }
          var arrayResult = this.duplicateCheck(checkArray);
          errorCode += arrayResult.errorCode;
          errorMsg += arrayResult.errorMsg;
    }
    errorMsg = 'Duplicated identifier:' + errorMsg;
    if (errorCode != 0)
      return errorMsg;
    return 0;
  }

  duplicateCheck(array) {
    var errorCode = 0;
    var errorMsg = '';
    for (var i = 0; i < (array.length-1); i++) {
      for (var j = i + 1; j < array.length; j++) {
        if (array[i] == array[j]  && array[i] !='__skip__') {
          errorCode += 1;
          errorMsg += array[i] + ' , ';
        }
      }
    }
    return {'errorCode' : errorCode , 'errorMsg' : errorMsg ,};
  }

  getDataByRow(sheet,idSeries,idRow,dataRow) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const numRow = this.getNumByRow(sheet,idSeries,idRow);
    const keysArray = Object.keys(numRow);
    var valueRow = {};
    for (var countKey = 0; countKey < keysArray.length; countKey++) {
      try {
        var isActive = this[keysArray[countKey]].active;} 
      catch(e) {
        var isActive = 0;}
      if (isActive == 1) 
        valueRow[keysArray[countKey]] = values[dataRow][numRow[keysArray[countKey]]].toString();
    }
    return valueRow;
  }

  getDataByCol(sheet,idSeries,idCol,dataCol) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const numCol = this.getNumByCol(sheet,idSeries,idCol);
    const keysArray = Object.keys(numCol);
    var valueCol = {};
    for (var countKey = 0; countKey < keysArray.length; countKey++) {
      try {
        var isActive = this[keysArray[countKey]].active;} 
      catch(e) {
        var isActive = 0;}
      if (isActive == 1) 
        valueCol[keysArray[countKey]] = values[numCol[keysArray[countKey]]][dataCol].toString();
    }
    return valueCol;
  }

}
