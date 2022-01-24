import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-import',
  templateUrl: './import.component.html',
  styleUrls: ['./import.component.css']
})
export class ImportComponent implements OnInit {

  rowArr :any[][];
  colArr: any[][];
  sourceArr : any[][];
  destArr : any[][];
  numRowLabels : number;
  numColumnLabels : number;
  rowSize :number;
  colSize :number;
  flattenData :any[][];
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor() { }

  ngOnInit(): void {
  }

  async uploadFile(){
        
    var myFile =<HTMLInputElement>document.getElementById("file");
    var reader = new FileReader();

  try{
    reader.onload = ( async (event) => {
      await Excel.run( async (context)=> {
            // Remove the metadata before the base64-encoded string.
            var startIndex = reader.result.toString().indexOf("base64,");
            var externalWorkbook = reader.result.toString().substr(startIndex + 7);

            var options = { 
              sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
              positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
              relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
          }; 

          var workbook = context.workbook;
          workbook.insertWorksheetsFromBase64(externalWorkbook, options);

          return context.sync();

        }).catch((err)=>{
          console.log(err);
          console.log("Unable to include the file");
        });
    });

    this.errorFlag = false;
    reader.readAsDataURL(myFile.files[0]); 
   }
    catch(err){
      console.log(err);
      this.errorFlag = true;
      this.errorMessage ="Unable to include the file";
    }

  }

  async deleteTable(){
    Excel.run( async (ctx)=>{
      var ws = ctx.workbook.worksheets.getItem("Data");
      ws.delete();
      this.errorFlag = false;
    }).catch((err)=>{
      console.log(err);
      console.log("No such file exists");
      this.errorFlag = true;
      this.errorMessage = "No such file exists";
    });
  }

  async toTable(){

    Excel.run(async (ctx)=>{
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getUsedRange();
      
      range.load();
      return ctx.sync().then(()=>{
        var selectedRange = range.address;
        var table = sheet.tables.add(selectedRange,true);
        table.name = 'sampleTable'
        return ctx.sync();
      } );
    
    }).catch((err)=>{
      console.log(err);
      console.log("Error in converting to table");
      this.errorFlag = true;
      this.errorMessage = err;
    })
  }

  toRange(){

    Excel.run(async (ctx)=>{
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getUsedRange();
      
      range.load();
      return ctx.sync().then(()=>{
        var selectedRange = range.address;
        var table = sheet.tables.getItemAt(0);
        table.getDataBodyRange().format.autofitColumns();
        // table.getRange().format.fill.color= '#ffffff';
        table.getRange().format.fill.clear();
        table.convertToRange();
        return ctx.sync();
      } );
    
    }).catch((err)=>{
      console.log(err);
      console.log("Error in converting to range");
      this.errorFlag = true;
      this.errorMessage = err;
    })

  }



}



