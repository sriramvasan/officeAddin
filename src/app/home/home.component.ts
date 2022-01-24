import { Component, ElementRef, OnInit, Output, ViewChild } from '@angular/core';

@Component({
  selector: 'home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  expensesTable : Excel.Table | undefined;
  isCreated : boolean ;
  tableNames :string[] = ["Orders"];
  sheetNames : string[] = ["Tables","Pivot"];
  option :number;
  example:any[];
  @ViewChild('inputValue')
  inputValue!: ElementRef;
  errorMessage : string ="Error";
  errorFlag : boolean = false;

  constructor(
    ) { 

  }

  ngOnInit(): void {
    this.getHeader();
  }
 
  async color(){
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRanges();
        range.load("address");
        range.format.fill.color = "magenta";
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }

  }

  async clearColor(){
    try{
      await Excel.run(async (content)=>{
        const range = content.workbook.getSelectedRanges();
        range.load('address');
        range.format.fill.clear();
        this.errorFlag = false;
        await content.sync();
      });
    }
    catch(error){
      console.log(error);
      this.errorFlag = true;
      this.errorMessage = "Error in coloring the range";
    }
  }

  async createTable(){
      await Excel.run( async (context)=>{

        var currentWorksheet = context.workbook.worksheets.add(this.sheetNames[0]);
        currentWorksheet.activate();
        var selectedRange = "A1:G1";
        this.expensesTable = currentWorksheet.tables.add(selectedRange,true);
        this.expensesTable.name = this.tableNames[0];
       
        this.expensesTable.getHeaderRowRange().values =
        [["Date", "Region", "Rep", "Item","Unit", "UnitCost","Total"]];

        this.expensesTable.showTotals = true;

     

    this.expensesTable.rows.add(undefined /*add at the end*/, [
        
        ["06-01-2020", "East", "Jones", "Pencil", 95, 1.99 , 189.05 ], 
        ["23-01-2020", "Central", "Kivell", "Binder", 50, 19.99 , 999.50 ], 
        ["09-02-2020", "Central", "Jardine", "Pencil", 36, 4.99 , 179.64 ], 
        ["26-02-2020", "Central", "Gill", "Pen", 27, 19.99 , 539.73 ], 
        ["15-03-2020", "West", "Sorvino", "Pencil", 56, 2.99 , 167.44 ], 
        ["01-04-2020", "East", "Jones", "Binder", 60, 4.99 , 299.40 ], 
        ["18-04-2020", "Central", "Andrews", "Pencil", 75, 1.99 , 149.25 ], 
        ["05-05-2020", "Central", "Jardine", "Pencil", 90, 4.99 , 449.10 ], 
        ["22-05-2020", "West", "Thompson", "Pencil", 32, 1.99 , 63.68 ], 
        ["08-06-2020", "East", "Jones", "Binder", 60, 8.99 , 539.40 ], 
        ["25-06-2020", "Central", "Morgan", "Pencil", 90, 4.99 , 449.10 ], 
        ["12-07-2020", "East", "Howard", "Binder", 29, 1.99 , 57.71 ], 
        ["29-07-2020", "East", "Parent", "Binder", 81, 19.99 , "1,619.19 "], 
        ["15-08-2020", "East", "Jones", "Pencil", 35, 4.99 , 174.65 ], 
        ["01-09-2020", "Central", "Smith", "Desk", 2, 125.00 , 250.00 ], 
        ["18-09-2020", "East", "Jones", "Pen Set", 16, 15.99 , 255.84 ], 
        ["05-10-2020", "Central", "Morgan", "Binder", 28, 8.99 , 251.72 ], 
        ["22-10-2020", "East", "Jones", "Pen", 64, 8.99 , 575.36 ], 
        ["08-11-2020", "East", "Parent", "Pen", 15, 19.99 , 299.85 ], 
        ["25-11-2020", "Central", "Kivell", "Pen Set", 96, 4.99 , 479.04 ], 
        ["12-12-2020", "Central", "Smith", "Pencil", 67, 1.29 , 86.43 ], 
        ["29-12-2020", "East", "Parent", "Pen Set", 74, 15.99 , "1,183.26 "], 
        ["15-01-2021", "Central", "Gill", "Binder", 46, 8.99 , 413.54 ], 
        ["01-02-2021", "Central", "Smith", "Binder", 87, 15.00 , "1,305.00 "], 
        ["18-02-2021", "East", "Jones", "Binder", 4, 4.99 , 19.96 ], 
        ["07-03-2021", "West", "Sorvino", "Binder", 7, 19.99 , 139.93 ], 
        ["24-03-2021", "Central", "Jardine", "Pen Set", 50, 4.99 , 249.50 ], 
        ["10-04-2021", "Central", "Andrews", "Pencil", 66, 1.99 , 131.34 ], 
        ["27-04-2021", "East", "Howard", "Pen", 96, 4.99 , 479.04 ], 
        ["14-05-2021", "Central", "Gill", "Pencil", 53, 1.29 , 68.37 ], 
        ["31-05-2021", "Central", "Gill", "Binder", 80, 8.99 , 719.20 ], 
        ["17-06-2021", "Central", "Kivell", "Desk", 5, 125.00 , 625.00 ], 
        ["04-07-2021", "East", "Jones", "Pen Set", 62, 4.99 , 309.38 ], 
        ["21-07-2021", "Central", "Morgan", "Pen Set", 55, 12.49 , 686.95 ], 
        ["07-08-2021", "Central", "Kivell", "Pen Set", 42, 23.95 , "1,005.90 "], 
        ["24-08-2021", "West", "Sorvino", "Desk", 3, 275.00 , 825.00 ], 
        ["10-09-2021", "Central", "Gill", "Pencil", 7, 1.29 , 9.03 ], 
        ["27-09-2021", "West", "Sorvino", "Pen", 76, 1.99 , 151.24 ], 
        ["14-10-2021", "West", "Thompson", "Binder", 57, 19.99 , "1,139.43 "], 
        ["31-10-2021", "Central", "Andrews", "Pencil", 14, 1.29 , 18.06 ], 
        ["17-11-2021", "Central", "Jardine", "Binder", 11, 4.99 , 54.89 ], 
        ["04-12-2021", "Central", "Jardine", "Binder", 94, 19.99 , "1,879.06 "], 
        ["21-12-2021", "Central", "Andrews", "Binder", 28, 4.99 , 139.7]
      ]);

        this.expensesTable.getRange().format.autofitColumns();
        this.expensesTable.getRange().format.autofitRows();
        await this.getHeader();
        this.isCreated = true;
        this.errorFlag = false;
    return context.sync();

      }).catch(()=>{
        console.log("Unable to load the table or already exists");
        this.errorFlag = true;
        this.errorMessage = "Unable to load the table or already exists";
        // this.alertMessageService.showErrorMessage(this.errorMessage);

      });
    
  }

  async clearTable(){
      await Excel.run(async(ctx)=>{
        ctx.workbook.worksheets.getItem(this.sheetNames[0]).delete();
        this.isCreated = false;  
        ctx.sync();
        }).catch(()=>{
          console.log("Table does not exist");
          this.errorFlag = true;
        this.errorMessage = "Table does not exist"
        });
   
  }

  async enterValue(){
    
      await Excel.run( async (context)=>{
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let selectedRange = context.workbook.getSelectedRange();
        let array:string = this.inputValue.nativeElement.value; 
        array = array.trim();
        selectedRange.values = [[array]];
        selectedRange.format.autofitColumns();
        this.errorFlag = false;
        return context.sync()
    }).catch(()=>{
      console.log("Some error occured while entering the value");
      this.errorFlag = true;
        this.errorMessage = "Unable to enter the value"
    });
  }

  async optionChoose(){
    var val= document.getElementById("name") as HTMLSelectElement;
    this.option = val.selectedIndex;
  }

  async getHeader(){
    Excel.run( async (ctx)=>{
    var worksheet = ctx.workbook.worksheets.getItem(this.sheetNames[0]);
    var headerrange = worksheet.tables.getItem(this.tableNames[0]).getHeaderRowRange().load("values");

    return ctx.sync().then(
      ()=>{
        this.example = headerrange.values[0];
        return ctx.sync();
      }
    );
  }).catch(()=>{
    console.log("Headers not present");
  });
  
  }

  SortDesc(){
    this.optionChoose();
    Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getItem(this.sheetNames[0]);
      var expensesTable = sheet.tables.getItem(this.tableNames[0]);
      
      var sortRange = expensesTable.getDataBodyRange();
      sortRange.sort.apply([
          {
              key: this.option,
              ascending: false,
          },
      ]);
      this.errorFlag = false;
      return context.sync();
  }).catch(()=>{
    console.log("Unable to Sort");
    this.errorFlag = true;
        this.errorMessage = "Unable to Sort"
  });
  }

  SortAsc(){
    this.optionChoose();
    Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getItem(this.sheetNames[0]);
      var expensesTable = sheet.tables.getItem(this.tableNames[0]);
      
      var sortRange = expensesTable.getDataBodyRange();
      sortRange.sort.apply([
          {
              key: this.option,
              ascending: true,
          },
      ]);
      this.errorFlag = false;
      return context.sync();
  }).catch(()=>{
    console.log("Unable to Sort");
    this.errorFlag = true;
        this.errorMessage = "Unable to Sort"
  });
  }
  
}
