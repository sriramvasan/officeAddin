import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { MessagesService } from '../messages.service';

@Component({
  selector: 'transform',
  templateUrl: './transform.component.html',
  styleUrls: ['./transform.component.css']
})
export class TransformComponent implements OnInit {

  @ViewChild('inputValue')
  inputValue!: ElementRef;
  tableNames :string[] = ["Orders","Sales"];
  sheetNames : string[] = ["Tables","Pivot"];
  isCreated : boolean =this.messageService.getTableStatus();
  example:any[];
  filterColumn:string;
  filterValue:string;
  filterType:string ;
  filterOption:number ;
  rangeAddress: string;
  selectedTableName : string;
  pivotViewOption : number = 0;
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor(private messageService: MessagesService) { }

  ngOnInit(): void {
    
    this.messageService.tableStatusChanged.subscribe(
      (status:boolean)=>{
        this.isCreated = status;
      })

    document.getElementById('pivotView').addEventListener('change',()=>{
      this.setPivotViewCustom();
    })
    
  }

  async createPivot(){
      Excel.run(async(ctx)=>{
        var rangeToAnalyze = ctx.workbook.worksheets.getItem(this.sheetNames[0]).tables.getItem(this.tableNames[0]);
        var currWorksheet = ctx.workbook.worksheets.add(this.sheetNames[1]);
        currWorksheet.activate();
        var rangeToPlacePivot = ctx.workbook.worksheets.getItem(this.sheetNames[1]).getRange("A1");
        ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.add(this.tableNames[1],rangeToAnalyze,rangeToPlacePivot);
  
        var pivotTable = ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.getItem(this.tableNames[1]);
        
        //// formatting options
        // pivotTable.layout.layoutType = "Tabular";
        // pivotTable.layout.repeatAllItemLabels(true);
        // pivotTable.layout.subtotalLocation = "Off";
        // pivotTable.layout.showRowGrandTotals = false;
        // pivotTable.layout.showColumnGrandTotals = false;

        // pivotTable.layout.showFieldHeaders = false;
        // // pivotTable.layout.emptyCellText = "0";
        // pivotTable.layout.fillEmptyCells = false;
       
        if(this.pivotViewOption ==0){
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Region"));
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Rep"));
  
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Item"));
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Unit"));
        }
        else if(this.pivotViewOption == 1){
          // pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Region"));
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Rep"));
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Item"));
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("UnitCost"));
        }
      
       
        pivotTable.dataHierarchies.load("no-properties-needed");
        pivotTable.layout.getRange().format.autofitColumns();
        pivotTable.layout.getRange().format.autofitRows();
        
        // pivotTable.dataHierarchies.remove(pivotTable.dataHierarchies.getItemOrNullObject("GrandTotal"))
        this.errorFlag = false;
        return ctx.sync();
      }).catch(()=>{
        this.errorFlag = true;
        console.log("Unable to find table or pivot already exists");
        this.errorMessage = "Unable to find table or pivot already exists";
      });
  }

  async deletePivot(){
      await Excel.run(async(ctx)=>{
        ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.getItem(this.tableNames[1]).delete();
        await ctx.sync();
        ctx.workbook.worksheets.getItem(this.sheetNames[1]).delete();
        this.errorFlag = false;
      }).catch(()=>{
        console.log("Pivot table does not exist");
        this.errorFlag = true;
        this.errorMessage = "Pivot table does not exist";
      });
    
  }

  getRangeAddress()
  /** 
   *  Gets the table selected in the selected range 
   */{
   Excel.run(async (ctx)=>{
      var range = ctx.workbook.getSelectedRange();
      var tables = range.getTables().load();
      range.load()
      return ctx.sync().then(()=>{
        this.rangeAddress = null;
        this.selectedTableName =  tables.items[0].name;
        var tblrng = tables.getItem(this.selectedTableName).getRange().load();
        
        this.errorFlag = false;
        return ctx.sync().then( ()=>{
          this.rangeAddress = this.selectedTableName+":("+tblrng.address+")";
        });
      });
    }).catch((err)=>{
      console.log(err);
      this.selectedTableName = null;
      console.log("Range does not have any table");
      this.errorFlag = true;
      this.errorMessage = "Range does not have any table";
    });
  }

  clearRange(){
   
    Excel.run(async (ctx)=>{

      if(this.selectedTableName!= null){
        ctx.workbook.tables.getItem(this.selectedTableName).getDataBodyRange().delete(Excel.DeleteShiftDirection.up);
      }
     
      // var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      // var selRange = sheet.getUsedRange().load();
      // var table = ctx.workbook.tables.load();
      // var tblnm = [];
      // var addresses = [];
      // var tblrange = [];

      // document.getElementById('same').innerHTML = "success";

      // return ctx.sync().then(()=>{
      //   table.items.forEach( element =>{
      //     tblnm.push(element.name);
      //     var r  = ctx.workbook.tables.getItem(element.name).getRange().load()
      //     tblrange.push(r)
      //     document.getElementById('same').innerHTML += "names1";
      //     document.getElementById('same').innerHTML += "<hr>" + tblnm.toString();
      //   });

      
      //   return ctx.sync().then(()=>{
      //     for(var i=0;i<tblrange.length;i++)
      //     {
      //       if(tblrange[i].address === selRange.address){
      //         document.getElementById('same').innerHTML += "<hr>" +"Table is present <br>"+tblrange[i].address;
      //       }
      //     }

          
      //     });

      //   });

      this.errorFlag = false;
      }).catch((err)=>{
        console.log(err);
        console.log("Unable to clear range");
        this.errorFlag = true;
        this.errorMessage = "Unable to clear range";
      });

  }

  filtering(){
    Excel.run(async (ctx)=>{
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var table = sheet.tables.getItemAt(0);
      var filter = table.columns.getItem(this.filterColumn).filter;

      var filtervalues =  this.filterValue.trim().split(',');
      
      for(var i =0;i<filtervalues.length;i++){
        filtervalues[i] = filtervalues[i].trim();
      }

      if(this.filterType === '0'){

        filter.apply({
          filterOn:Excel.FilterOn.values,
          // values:[this.filterValue.trim()]
          values :filtervalues
        });
      }
      // document.getElementById('same').innerHTML += "success1"+"  filter type chosen "+this.filterType.toString()+" "+typeof(this.filterType);;
      if(this.filterType === '1'){
        filter.apply({
          criterion1 : this.filterOption+this.filterValue ,
          filterOn : Excel.FilterOn.custom
        })
      }
      this.errorFlag = false;
      return ctx.sync();
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to filter");
      this.errorFlag = true;
      this.errorMessage = "Unable to filter";
    });
  }

  clearFilter(){
    Excel.run(async (ctx)=>{
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var table = sheet.tables.getItemAt(0);
      // var filter = table.columns.getItem(this.filterColumn).filter;
      table.clearFilters();
      var date = Date.now();
      // document.getElementById('same').innerHTML = new Date(date).toDateString()+
      //           "<hr>"+ new Date(date).toTimeString();
    }).catch((err)=>{
      console.log(err);
      console.log("Unable to clear filter")
    });
    this.filterColumn = null;
    this.filterOption = null;
    this.filterType = null;
    this.filterValue = null;
  }

  deleteSelectedTable(){
    Excel.run(async(ctx)=>{

    //  var tables =  ctx.workbook.tables.load(); 
    //   var tables = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange(); // getTables();
      var tables = ctx.workbook.tables;

      tables.load();
      tables.getItemAt(0).getTotalRowRange().getEntireColumn().load()
      // var range = tables.items[0].getRange().load()
    tables.load();

    this.errorFlag = false;

     return ctx.sync().then(()=>{
     document.getElementById('same').innerHTML += tables.items[0].name;
    //  document.getElementById('same').innerHTML += range.address;
     })
     
    }).catch((err)=>{
      console.log(err);
      console.log("Cannot delete table");
      this.errorFlag = true;
      this.errorMessage = "Cannot delete table/ no table";
    });
  }


  async setPivotViewCustom(){

    Excel.run(async (ctx)=>{

      var pivotTable = ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.getItem(this.tableNames[1]);
      // var select = document.getElementById("pivotView") as HTMLSelectElement;
      var rangeToAnalyze = ctx.workbook.worksheets.getItem(this.sheetNames[0]).tables.getItem(this.tableNames[0]);
      var rangeToPlacePivot = ctx.workbook.worksheets.getItem(this.sheetNames[1]).getRange("A1");


      if(this.pivotViewOption == 0)
        {
          pivotTable.delete();
          ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.add(this.tableNames[1],rangeToAnalyze,rangeToPlacePivot);
          pivotTable = ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.getItem(this.tableNames[1]);

          

          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Region"));
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Rep"));
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Item"));
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Unit"));
        }
        else if(this.pivotViewOption == 1 )
        {

          pivotTable.delete();
          ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.add(this.tableNames[1],rangeToAnalyze,rangeToPlacePivot);
          pivotTable = ctx.workbook.worksheets.getItem(this.sheetNames[1]).pivotTables.getItem(this.tableNames[1]);

          // pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Region"));
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Rep"));
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Item"));
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("UnitCost"));
        }

        this.errorFlag = false;
    }).catch((err)=>{
      console.log(err);
      this.errorFlag = true;
      this.errorMessage = err;
    })

  }  

}