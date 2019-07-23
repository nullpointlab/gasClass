/**
* @auther nosuke
* @update 2019-07-07
*/
function sheetOperation(sheetId) {
  //コンストラクタ  
  this.ss = SpreadsheetApp.openById(sheetId);    
  this.values = this.ss.getDataRange().getValues();

  /**
  * @description 行列の始まりを取得します。 
  * @param {String} "column"か"row"を引数に入れてください。
  * @return {Integer}
  */
  this.getFstNum = function(matrix){
    for (var i = 0; i < this.values.length; i++) {
      for (var j = 0; j < this.values[i].length; j++) {
        if(!this.values[i][j]){
          continue;
        }else{
          return (matrix=="column")?i:j;
        }
      }
    }
  }
  
  /**
  * @description 列の終わりを取得します。 
  * @return {Integer}
  */  
  this.getLastColNum = function(){
    return this.values.length;
  }
  
  /**
  * @description 行の終わりを取得します。 
  * @return {Integer}
  */    
  this.getLastRowNum = function(column){
    return this.values[column].length;
  }
  
  /**
  * @description 対象シートから文字列を検索します。 
  * @param {String} 検索文字を引数に入れてください。
  * @return {Array} 行と列を配列で返却します。
  */    
  this.findString = function(str){
    for (var i = 0; i < this.values.length; i++) {
      for (var j = 0; j < this.values[i].length; j++) {
        if(this.values[i][j]==str){
          return [i,j];
        }
      }
    }
  }
  
  /**
  * @description 対象シートの指定セルにデータをセットします。 
  * @param {int}　開始行
  * @param {int}　開始列
  * @param {int}　終了行
  * @param {int}　終了列
  * @param {Array} 指定セル同様の配列を作成し、引数に入れてください。
  */  
  this.setDataValues = function(startRow,startCol,endRow,endCol,value){
    this.ss.getSheets()[0].getRange(startRow,startCol,endRow,endCol).setValues(value);
  }
  
  /**
  * @description 最終行に1列目からデータをセットします。
  * @param {Array} 2次元配列でセットしてください。
  */      
  this.setDataLstRowValues = function(value){
    this.ss.getSheets()[0].getRange(this.values[0].length,1,value.length,value[0].length).setValues(value);
  }
  
  /**
  * @description 対象シートから文字列を検索します。 
  * @param {String} 検索文字を引数に入れてください。
  */     
  this.setDataLstColValues = function(value){
    this.ss.getSheets()[0].getRange(1,this.values[0].length+1,value.length,value[0].length).setValues(value);
  }
  
}


