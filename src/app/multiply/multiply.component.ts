import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-multiply',
  templateUrl: './multiply.component.html',
  styleUrls: ['./multiply.component.css']
})
export class MultiplyComponent implements OnInit {


  range1 : Excel.Range;
  range2 : Excel.Range;
  range1Array : any[][];
  range2Array : any[][];
  currStepIndex :number = 0;
  acquiredArray: string;
  isLoader : boolean = true;
  destinationRangeAddress: string;
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor() { }

  ngOnInit(): void {
  }

  multiplyRanges(){
  
    Excel.run(async (ctx)=>{

      if((this.range1Array== null) || (this.range2Array == null)){
        this.errorFlag = true;
        console.log("Not enough arrays chosen to multiply");
        this.errorMessage = "Not enough arrays chosen to multiply"
      }
      else{ 
        var val;
        // val= [[1,2,3],[1,2,3],[1,2,3]]
        val = MultiplyComponent.multiplyRangesinArray(this.range1Array,this.range2Array);
        // console.log(this.range1Array);
        // console.log(this.range2Array);
        var numCols = val[0].length;
        var numRows = val.length;
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        var destRng = sheet.getRange(this.destinationRangeAddress);
        destRng= destRng.getResizedRange(numRows-1,numCols-1);
        destRng.values = val;
        this.errorFlag = false;
      }
    return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to multiply ranges");
      this.errorFlag = true;
      this.errorMessage = "Unable to multiply ranges";
    });

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

  async getRange(){
    if(this.currStepIndex==0){
      await this.getRangeFirst();
      console.log("first time getting array");
    }
    else if(this.currStepIndex >=1){
      await this.getRangeAgain();
      console.log("not the first time");
      try{
        this.range1Array = MultiplyComponent.multiplyRangesinArray(this.range1Array,this.range2Array);
        this.errorFlag = false;
      }
      catch(err){
        console.log(err);
        this.errorFlag = true;
        this.errorMessage = err;
      }
    }
    
    
  }

  nextRange(){
    this.currStepIndex ++;
    this.acquiredArray = null;
  }

  async getRangeFirst(){
    Excel.run( async (ctx)=>{
        this.range1 = ctx.workbook.getSelectedRange();
        this.range1 = this.range1.getUsedRange();
        this.range1.load();

        this.errorFlag = false;
        return ctx.sync().then(()=>{
          this.acquiredArray = this.range1.address;
          this.range1Array  = this.range1.values;
          console.log(this.range1Array);
        });
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get first range");
      this.errorFlag = true;
      this.errorMessage = "Unable to get first range";
    }); 
  }

  async getRangeAgain(){
    Excel.run( async(ctx)=>{
      this.range2 = ctx.workbook.getSelectedRange();
      this.range2 = this.range2.getUsedRange();
      this.range2.load();

      this.errorFlag = false;
      return ctx.sync().then(()=>{
        this.acquiredArray = this.range2.address;
        this.range2Array = this.range2.values;
        console.log(this.range2Array);
      });
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get range",this.currStepIndex+1);
      this.errorFlag = true;
      this.errorMessage = "Unable to get range";
    });
  }

  async cancelLoad(){
    this.isLoader = !this.isLoader;
  }

  getDestRange(){
    Excel.run(async (ctx)=>{
      var destRng = ctx.workbook.getSelectedRange();
      destRng.load();
      this.errorFlag = false;
      return ctx.sync().then(()=>{
        this.destinationRangeAddress = destRng.address 
      });
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to get destination");
      this.errorFlag = true;
      this.errorMessage = "Unable to get destination";
    });

  }

  reset(){
    this.range1Array = null;
    this.range2Array = null;
    this.currStepIndex = 0;
    this.destinationRangeAddress = null;
    this.acquiredArray = null;
  }


}
