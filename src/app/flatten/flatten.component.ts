import { Component, OnInit } from '@angular/core';

export enum ReturnType{
  Address = 0,
  Range= 1,
  Array = 2

}

@Component({
  selector: 'app-flatten',
  templateUrl: './flatten.component.html',
  styleUrls: ['./flatten.component.css']
})
export class FlattenComponent implements OnInit {

    flatteningStep :number =0;

  rowLabelRange : Excel.Range|undefined=undefined;
  rowLabelAddress :string|undefined= undefined;

  columnLabelRange: Excel.Range |undefined = undefined;
  columnLabelAddress :string |undefined =undefined;

  dataRange:Excel.Range |undefined =undefined;
  dataRangeAddress : string|undefined =undefined;

  destinationRange: Excel.Range|undefined =undefined;
  destinationRangeAddress: string|undefined =undefined;

  //all datas
  rowArr :any[][]=[];
  colArr: any[][] = [];
  sourceArr : any[][]= [];
  destArr : any[][] =[];

  numRowLabels : number =-1;
  numColumnLabels : number =-1;
  rowSize :number = -1;
  colSize :number =-1;
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor() { }

  ngOnInit(): void {
  }

  startFlatten(){
    this.resetFlatten();
    this.flatteningStep++;
  }

  async getRowLabels(){
    
   await Excel.run(async (ctx)=>{
     let sheet = ctx.workbook.worksheets.getActiveWorksheet();
     if(this.rowLabelAddress == undefined){
       this.rowLabelRange = ctx.workbook.getSelectedRange();
     }
     else{
       this.rowLabelRange = sheet.getRange(this.rowLabelAddress);
     }
      this.rowLabelRange = this.rowLabelRange.getUsedRange();
      this.rowLabelRange.load()
      await ctx.sync();

      this.rowLabelAddress = (<Excel.Range>this.rowLabelRange).address
      this.errorFlag = false;
      console.log(this.rowLabelAddress);

      return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      this.errorFlag= true;
      this.errorMessage = err;
    });
  }

  async getColumnLabels(){
    await Excel.run(async (ctx)=>{
      let sheet = ctx.workbook.worksheets.getActiveWorksheet();

      if(this.columnLabelAddress == undefined){
        this.columnLabelRange = ctx.workbook.getSelectedRange();
      }
      else{
        this.columnLabelRange = sheet.getRange(this.columnLabelAddress);
      }
      this.columnLabelRange = this.columnLabelRange.getUsedRange();
      this.columnLabelRange.load()
      await ctx.sync();
      this.columnLabelAddress = (<Excel.Range>this.columnLabelRange).address 
      this.errorFlag = false;
      console.log(this.columnLabelAddress);

      return ctx.sync();
      
    }).catch((err)=>{
      console.log(err);
      this.errorFlag= true;
      this.errorMessage = err;
    });

  }

  async getDestRng(){

    await Excel.run(async (ctx)=>{
      let sheet = ctx.workbook.worksheets.getActiveWorksheet();
      let destRng;
     if( this.destinationRangeAddress== undefined){
       this.destinationRange = ctx.workbook.getSelectedRange();
       destRng = ctx.workbook.getSelectedRange();
     }
     else{
       this.destinationRange = sheet.getRange(this.destinationRangeAddress);
     }
      this.destinationRange = this.destinationRange.getAbsoluteResizedRange(1,1);
      this.destinationRange.load()
      await ctx.sync();
      this.destinationRangeAddress = (<Excel.Range>this.destinationRange).address 
      this.errorFlag = false;
      console.log(this.destinationRangeAddress);
      
      return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      this.errorFlag= true;
      this.errorMessage = err;
    });
  }

  async getSourceData(){

    await Excel.run( async (ctx)=>{
      let dataSheet = ctx.workbook.worksheets.getActiveWorksheet();

      let colLblRng = dataSheet.getRange(this.columnLabelAddress).load();
      let rowLblRng = dataSheet.getRange(this.rowLabelAddress).load();
      this.dataRange = colLblRng.getEntireColumn().getUsedRange().getIntersection(rowLblRng.getEntireRow().getUsedRange()).load('values');;
      this.dataRange.load();
      await ctx.sync();

      (<Excel.Range>this.dataRange).select()
      this.dataRangeAddress = (<Excel.Range>this.dataRange).address;
      this.colArr = colLblRng.values;
      this.rowArr = rowLblRng.values;
      this.sourceArr = (<Excel.Range>this.dataRange).values;
      this.numRowLabels = rowLblRng.columnCount;
      this.numColumnLabels = colLblRng.rowCount;
      this.rowSize = (<Excel.Range>this.dataRange).rowCount;
      this.colSize = (<Excel.Range>this.dataRange).columnCount;
      this.errorFlag = false;
      console.log(this.dataRangeAddress);
      
      return ctx.sync();
      
    }).catch((err)=>{
      console.log(err);
      this.errorFlag= true;
      this.errorMessage = err;
    });
  }

  async embed(){

    await Excel.run(async (ctx)=>{

      let newcolscount = this.numColumnLabels+this.numRowLabels+1;
      let newrowscount = this.colSize* this.rowSize;
    
      let rowLblPosition = 0
      let colLblPosition = this.numRowLabels;
      let flatSheet = ctx.workbook.worksheets.getActiveWorksheet();     
      let destinationArray = new Array(newcolscount);
      let destRng2 = flatSheet.getRange(this.destinationRangeAddress);

      destRng2 = destRng2.getResizedRange(newrowscount-1,newcolscount-1);
      
      let sampleMatrix = []
      for(let i=0 ;i<newrowscount;i++){
        sampleMatrix[i] = new Array(newcolscount);
      }

      let samp=0;
    
      for(let colIndex =0; colIndex<this.colArr[0].length ;colIndex++)
      {
        for(let rowIndex = 0; rowIndex < this.rowArr.length ; rowIndex++){

          for( let rowLabelIndex =0 ; rowLabelIndex < this.rowArr[0].length;rowLabelIndex++){
            sampleMatrix[samp][rowLblPosition+rowLabelIndex] = this.rowArr[rowIndex][rowLabelIndex];
          }
          
          for(let colLabelIndex =0; colLabelIndex<this.colArr.length;colLabelIndex++){
            sampleMatrix[samp][colLblPosition+colLabelIndex] = this.colArr[colLabelIndex][colIndex];
          }

          sampleMatrix[samp][this.numColumnLabels+this.numRowLabels] = this.sourceArr[rowIndex][colIndex];
          samp++;  
        }
      }
      
      destRng2.values = sampleMatrix;
      this.errorFlag = false;
      return ctx.sync()
    }).catch((err)=>{
      console.log(err);
      this.errorFlag= true;
      this.errorMessage = err;
    });
  }

  async nextStep(){

    if(this.flatteningStep == 1){
      await this.getColumnLabels();
      console.log(this.columnLabelAddress);
      this.flatteningStep++;
    }
    else if(this.flatteningStep ==2){
      await this.getRowLabels();
      console.log(this.rowLabelAddress);
      this.getSourceData();
      this.flatteningStep++;
    }

    else if(this.flatteningStep ==3){
     this.getSourceData();
      console.log(this.sourceArr);
      this.flatteningStep++;
    }
    else if(this.flatteningStep ==4){
      console.time("Flatten");
      await this.getDestRng();
      await this.embed();
      console.timeEnd("Flatten");
      console.log(this.destinationRangeAddress);
      // this.flatteningStep = 0;
    }
    else{
      this.flatteningStep =0;
      return;
    }
  }

  resetFlatten(){
    this.flatteningStep = 0;
    this.columnLabelAddress = undefined;
    this.rowLabelAddress = undefined;
    this.dataRangeAddress = undefined;
    this.destinationRangeAddress = undefined;
  }

  static async getsourceData(rowLabelAddress?: string, colLabelAddress?:string,sourceSheetName:string|undefined = undefined ,returnType:ReturnType = ReturnType.Address){

    let dataRngAddress :string ='';
    let sourceArray :any[][] = [];
    let dataRng : Excel.Range|undefined = undefined;

    await Excel.run( async (ctx)=>{
      let dataSheet;
      if(sourceSheetName== undefined){

        dataSheet = ctx.workbook.worksheets.getActiveWorksheet();
      }else{
        dataSheet = ctx.workbook.worksheets.getItem(sourceSheetName);
      }

      
      let colLblRng = dataSheet.getRange(colLabelAddress).load();
      // colLblRng.select();
      let rowLblRng = dataSheet.getRange(rowLabelAddress).load();

      // dataRng = colLblRng.getEntireColumn().getUsedRange().getIntersection(rowLblRng.getEntireRow().getUsedRange()).load('values');
      dataRng = colLblRng.getEntireColumn().getUsedRange().getIntersection(rowLblRng.getEntireRow()).load('values');
      dataRng.load();
      
      await ctx.sync();

      dataRngAddress = dataRng.address;
      sourceArray = dataRng.values;
      return ctx.sync();
    })
    .catch((err)=>{
      console.log(err);
    });

    if(returnType==0) {
      return dataRngAddress;
    }
    else if(returnType ==1){ 
      return sourceArray;
    }
    else{
      return dataRng;
    }

  }

  static async Flatten(rowLblAddress :string , colLbladdress : string,destAddress:string,sourceSheetName:string|undefined = undefined,destSheetName:string|undefined = undefined ){
    
    let flag = false;

    await Excel.run(async ctx=>{
      let sheet;
      if(sourceSheetName == undefined){

        sheet = ctx.workbook.worksheets.getActiveWorksheet();
      }else{
        sheet = ctx.workbook.worksheets.getItem(sourceSheetName);
      }

      let colLblRng = sheet.getRange(colLbladdress);
      // colLblRng = colLblRng.getRowsAbove();

      let rowLblRng = sheet.getRange(rowLblAddress);

      let dataRngAddress = await FlattenComponent.getsourceData(rowLblAddress, colLbladdress,sourceSheetName,ReturnType.Address) as string;
      let dataRng = sheet.getRange(dataRngAddress);

      colLblRng.load(['rowCount','values']);
      rowLblRng.load(['columnCount','values']);
      dataRng.load(['rowCount','columnCount','values']);

      await ctx.sync();

      let numRowLbls = rowLblRng.columnCount;
      let numCollbls = colLblRng.rowCount;
      let rowSize = dataRng.rowCount;
      let colSize = dataRng.columnCount;

      let colArray = colLblRng.values;
      let rowArray = rowLblRng.values;
      let sourceArray = dataRng.values;

      // console.log(sourceArray);

      let newcolscount = numCollbls+numRowLbls+1;
      let newrowscount = colSize* rowSize;

      let rowLblPosition = 0
      let colLblPosition = numRowLbls;

      let flatSheet;
      if(destSheetName === undefined){
        flatSheet = ctx.workbook.worksheets.getActiveWorksheet();
      }
      else{
        flatSheet = ctx.workbook.worksheets.getItem(destSheetName);     
      }

      let destRng = flatSheet.getRange(destAddress);
      destRng = destRng.getResizedRange(newrowscount-1,newcolscount-1);

      let sampleMatrix = []
      for(let i=0 ;i<newrowscount;i++){
        sampleMatrix[i] = new Array(newcolscount);
      }

      let samp=0;
    
      for(let colIndex =0; colIndex<colArray[0].length ;colIndex++)
      {
        for(let rowIndex = 0; rowIndex < rowArray.length ; rowIndex++){

          for( let rowLabelIndex =0 ; rowLabelIndex < rowArray[0].length;rowLabelIndex++){
            sampleMatrix[samp][rowLblPosition+rowLabelIndex] = rowArray[rowIndex][rowLabelIndex];
          }
          
          for(let colLabelIndex =0; colLabelIndex<colArray.length;colLabelIndex++){
            sampleMatrix[samp][colLblPosition+colLabelIndex] = colArray[colLabelIndex][colIndex];
          }

          sampleMatrix[samp][numCollbls+numRowLbls] = sourceArray[rowIndex][colIndex];
          samp++;  
        }
      }

      destRng.values = sampleMatrix;
      flag = true;
      return ctx.sync();
    })
    .catch(err =>{
      console.log(err);

      flag = false;
    })

    return flag;
  }


}
