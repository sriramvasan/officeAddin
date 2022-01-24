import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-flatten',
  templateUrl: './flatten.component.html',
  styleUrls: ['./flatten.component.css']
})
export class FlattenComponent implements OnInit {


  flatteningStep :number =0;

  rowLabelRange : Excel.Range=null;
  rowLabelAddress :string=null;

  columnLabelRange: Excel.Range = null;
  columnLabelAddress :string=null;

  dataRange:Excel.Range = null;
  dataRangeAddress : string=null;

  destinationRange: Excel.Range=null;
  destinationRangeAddress: string=null;

  //all datas
  rowArr :any[][];
  colArr: any[][];
  sourceArr : any[][];
  destArr : any[][];

  numRowLabels : number;
  numColumnLabels : number;
  rowSize :number;
  colSize :number;
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor() { }

  ngOnInit(): void {
  }

  startFlatten(){
    this.flatteningStep++;
  }
 
  async getRowLabels(){
    
    Excel.run(async (ctx)=>{
      this.rowLabelRange = ctx.workbook.getSelectedRange();
      this.rowLabelRange = this.rowLabelRange.getUsedRange();
      this.rowLabelRange.load()

      return ctx.sync().then(()=>{
        this.rowLabelAddress = this.rowLabelRange.address 
        this.errorFlag = false;
      });
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get row labels");
      this.errorFlag =true;
      this.errorMessage = "Unable to get row labels";
    });
  }

  async getColumnLabels(){
    Excel.run(async (ctx)=>{
      this.columnLabelRange = ctx.workbook.getSelectedRange();
      this.columnLabelRange = this.columnLabelRange.getUsedRange();
      this.columnLabelRange.load()
      return ctx.sync().then(()=>{
        this.columnLabelAddress = this.columnLabelRange.address 
        this.errorFlag = false;
      });
      
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get column labels");
      this.errorFlag = true;
      this.errorMessage = "Unable to get column labels";
    });

  }

  async getsourceData(){

    Excel.run( async (ctx)=>{

      var dataSheet = ctx.workbook.worksheets.getItem('SampleData');
      var colLblRng = dataSheet.getRange(this.columnLabelAddress).load();
      var rowLblRng = dataSheet.getRange(this.rowLabelAddress).load();

      this.dataRange = colLblRng.getEntireColumn().getUsedRange().getIntersection(rowLblRng.getEntireRow().getUsedRange()).load('values');
      // this.dataRange = this.columnLabelRange.getEntireColumn().getUsedRange().getIntersection(this.rowLabelRange.getEntireRow().getUsedRange()).load();
      this.dataRange.load();
      
      return ctx.sync().then(()=>{
        this.dataRangeAddress = this.dataRange.address;
        this.colArr = colLblRng.values;
        this.rowArr = rowLblRng.values;
        this.sourceArr = this.dataRange.values;
        this.numRowLabels = rowLblRng.columnCount;
        this.numColumnLabels = colLblRng.rowCount;
        this.rowSize = this.dataRange.rowCount;
        this.colSize = this.dataRange.columnCount;
        this.errorFlag = false;
      });
      
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get source data");
      this.errorFlag = true;
      this.errorMessage = "Unable to get source data";
    });

  }

  async getDestRng(){

    Excel.run(async (ctx)=>{
      this.destinationRange = ctx.workbook.getSelectedRange();
      this.destinationRange.load()
      return ctx.sync().then(()=>{
        this.destinationRangeAddress = this.destinationRange.address 
        this.errorFlag = false;
      });
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get destination range");
      this.errorFlag = true;
      this.errorMessage = "Unable to get destination range";
    });
  }

  async embed(){
    
    Excel.run(async (ctx)=>{
      
      var newcolscount = this.numColumnLabels+this.numRowLabels+1;
      var newrowscount = this.colSize* this.rowSize;
    
      var rowLblPosition = 0
      var colLblPosition = this.numRowLabels;
      var flatSheet = ctx.workbook.worksheets.getItem("SampleData");      
      var destinationArray = new Array(newcolscount);
      var destRng2 = flatSheet.getRange(this.destinationRangeAddress);

      destRng2.format.fill.color = 'orange';
      destRng2 = destRng2.getResizedRange(newrowscount-1,newcolscount-1);
      // destRng2.format.fill.color = 'magenta';
      // destRng2.format.fill.color = "lightgreen";
      
      var sampleMatrix = []
      for(var i=0 ;i<newrowscount;i++){
        sampleMatrix[i] = new Array(newcolscount);
      }

      var samp=0;
    
      for(var colIndex =0; colIndex<this.colArr[0].length ;colIndex++)
      {
        for(var rowIndex = 0; rowIndex < this.rowArr.length ; rowIndex++){

          for( var rowLabelIndex =0 ; rowLabelIndex < this.rowArr[0].length;rowLabelIndex++){
            sampleMatrix[samp][rowLblPosition+rowLabelIndex] = this.rowArr[rowIndex][rowLabelIndex];
          }
          
          for(var colLabelIndex =0; colLabelIndex<this.colArr.length;colLabelIndex++){
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
      console.log("Unable to flatten");
      this.errorFlag = true;
      this.errorMessage = "Unable to flatten";
    });
  }
  
  resetFlatten(){
    this.flatteningStep = 0;
    this.columnLabelAddress = null;
    this.rowLabelAddress = null;
    this.dataRangeAddress = null;
    this.destinationRangeAddress = null;
  }

  nextStep(){
    if(this.flatteningStep ===4){
      this.embed();
      this.flatteningStep = 0
    }
    else if(this.flatteningStep ===2){
      this.getsourceData();
      this.flatteningStep++;
    }
    else 
    {
      this.flatteningStep++;}
    
  }
}
