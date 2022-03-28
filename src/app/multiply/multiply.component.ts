import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-multiply',
  templateUrl: './multiply.component.html',
  styleUrls: ['./multiply.component.css']
})
export class MultiplyComponent implements OnInit {

  rangeArray : any[][];
  errorMessage : string ="Error";
  errorFlag : boolean = false;
  selectedRanges :string[] = [];
  constructor() { }

  ngOnInit(): void {
  }

  static multiplyRangesinArray(array1:any[][] , array2:any[][]){
    var outputArray: any[][];
    outputArray = new Array(array1.length * array2.length);
    let index =0;
    for(var i =0;i< array1.length ;i++){
      for(var j=0;j< array2.length;j++){
        outputArray[index] = array1[i].concat(array2[j]);
        index++;
      }
    }
    return outputArray;
  }

  async getDestRange(){
    let destinationRangeAddress = "";
    await Excel.run(async (ctx)=>{
      var destRng = ctx.workbook.getSelectedRange();
      destRng.load();
      this.errorFlag = false;
      await ctx.sync();
      destinationRangeAddress = destRng.address;
      return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get destination");
      this.errorFlag = true;
      this.errorMessage = "Unable to get destination";
    });
    return destinationRangeAddress;
  }

  reset(){
    this.rangeArray = null;
    this.selectedRanges = [];
  }

  async getSelectedRange(){
    let selectedArray = [];
    await Excel.run( async (ctx)=>{
    
      let range = ctx.workbook.getSelectedRange();
      range = range.getUsedRange();
      range.load();

      this.errorFlag = false;
      await ctx.sync();
      selectedArray  = range.values;
      console.log(this.rangeArray);
      
    }).catch((err)=>{
    console.log(err);
    console.log("Unable to get range");
    this.errorFlag = true;
    this.errorMessage = "Unable to get range";
  }); 
  return selectedArray;
  }

  async LoadRanges(){
    let newArray = [];
    if(this.selectedRanges.length ==0){
      this.rangeArray = await this.getSelectedRange();
      console.log("first time getting array");
      this.selectedRanges.push("Range 1 :"+this.rangeArray.toString())
    }
    else if(this.selectedRanges.length >=1){
    newArray = await this.getSelectedRange();
      this.rangeArray = MultiplyComponent.multiplyRangesinArray(this.rangeArray,newArray);
      this.errorFlag = false;
      this.selectedRanges.push("Range "+ (this.selectedRanges.length+1) +":"+newArray) 
    }
    
  }

  async multipy(){
    let destinationRangeAddress = await this.getDestRange();
    let val = this.rangeArray;
    await Excel.run(async ctx=>{
      var numCols = val[0].length;
      var numRows = val.length;
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var destRng = sheet.getRange(destinationRangeAddress);
      destRng= destRng.getResizedRange(numRows-1,numCols-1);
      destRng.values = val;
      this.errorFlag = false;
      return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to multiply ranges");
      this.errorFlag = true;
      this.errorMessage = "Unable to multiply ranges";
    });
  }
}
