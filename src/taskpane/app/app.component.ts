import { analyzeAndValidateNgModules } from "@angular/compiler";
import { Component, OnInit } from "@angular/core";
import { stringify } from "@angular/core/src/util";
 
/* global console, Excel */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent   implements OnInit{
  welcomeMessage = "Welcome";
Data : any; 

constructor() {     
}

ngOnInit() { 
}
 
 async CreateTable() {
    try {
      
      await Excel.run(async (context) => {

        context.runtime.load("enableEvents");
        await context.sync();

        let eventBoolean = !context.runtime.enableEvents;
        context.runtime.enableEvents = eventBoolean;

        /**
         * Insert your Excel code here
         */
         let URL = "https://localhost:44339/api/values";
         let promise = new Promise(function (resolve, reject) {
             let req = new XMLHttpRequest();
             req.open("GET", URL);
             req.onload = function () {
                 if (req.status == 200) {
                     resolve(req.response);
                 } else {
                     reject("There is an Error!");
                 }
             };
             req.send();
         });
         promise.then(
             async (result) => {
               this.Data= result;

          var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
          
                    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                        currentWorksheet.getUsedRange().clear();
                        context.application.suspendApiCalculationUntilNextSync();

                     }

          var expensesTable = currentWorksheet.tables.add("A1:P1", true /*hasHeaders*/);
          expensesTable.name = "ExpensesTable";
           
      
          expensesTable.getHeaderRowRange().values =
                        [["ID", "FID", "SECURITYID", "SECURITYNAME", "SECURITYTYPE", "EXPOSURE", "SECURITYTYPEMAPPED", "EXPOSUREMAPPED", "OPENINGCOST", "PURCHASES", "SALES", "AMORTIZATION", "NETREALIZED", "CLOSINGMARKETVALUE", "EXCEPTIONS","RECONCILIATION"]];
       
                    var contact = JSON.parse(this.Data);

                    var newData = contact.map(item =>
                        [item.ID, item.FID, item.SECURITYID, item.SECURITYNAME, item.SECURITYTYPE, item.EXPOSURE, item.SECURITYTYPEMAPPED, item.EXPOSUREMAPPED, item.OPENINGCOST, item.PURCHASES, item.SALES, item.AMORTIZATION, item.NETREALIZED, item.CLOSINGMARKETVALUE, item.EXCEPTIONS, item.RECONCILIATION]);
                 
                        expensesTable.rows.add(null, newData);
                  
                        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                        currentWorksheet.getUsedRange().format.autofitColumns();
                        currentWorksheet.getUsedRange().format.autofitRows();

                      }
                      
                     var range = currentWorksheet.getRange("I2:K520");

                     range.dataValidation.rule = {
                         wholeNumber: {
                             formula1: 0,
                             operator: "GreaterThan"
                         }
                     };

                     var comments = currentWorksheet.comments;
                     comments.add("Sheet1!B3", "TODO: add data......");           
   
                     await  context.sync();
                    },
                    (error) => {
                        // A rejected prmise will execute this
                        console.log('We have encountered an Error!'); // Log an error
                    }
                );
               
              });
            } catch (error) {
              console.error(error);
            }
          }

  async DynaminCreateTable() {
    try {
      
      var e = (document.getElementById("fund")) as HTMLSelectElement;
      var sel = e.selectedIndex;
      var opt = e.options[sel];
      var strfund = opt.value;
      var e = (document.getElementById("routine")) as HTMLSelectElement;
      var sel = e.selectedIndex;
      var opt = e.options[sel];
      var strroutine = opt.value;
         
       
      
      await Excel.run(async (context) => {
       
        context.runtime.load("enableEvents");
        await context.sync();
        let eventBoolean = !context.runtime.enableEvents;
        context.runtime.enableEvents = eventBoolean;

        let URL = "https://localhost:44339/api/values/GetDynamic?efund=" + strfund + "&eroutine=" + strroutine;
         let promise = new Promise(function (resolve, reject) {
             let req = new XMLHttpRequest();
             req.open("GET", URL);
             req.onload = function () {
                 if (req.status == 200) {
                     resolve(req.response);
                 } else {
                     reject("There is an Error!");
                 }
             };
             req.send();
         });
         promise.then(
             async (result) => {
               this.Data= result;

          var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
                    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {

                         currentWorksheet.getUsedRange().clear();
                       context.application.suspendApiCalculationUntilNextSync();
                    }
          var expensesTable = currentWorksheet.tables.add("A1:P1", true /*hasHeaders*/);
          expensesTable.name = "ExpensesTable";
         
    expensesTable.getHeaderRowRange().values =
                        [["ID", "FID", "SECURITYID", "SECURITYNAME", "SECURITYTYPE", "EXPOSURE", "SECURITYTYPEMAPPED", "EXPOSUREMAPPED", "OPENINGCOST", "PURCHASES", "SALES", "AMORTIZATION", "NETREALIZED", "CLOSINGMARKETVALUE", "EXCEPTIONS","RECONCILIATION"]];

                   var contact = JSON.parse(this.Data);

                    var newData = contact.map(item =>
                        [item.ID, item.FID, item.SECURITYID, item.SECURITYNAME, item.SECURITYTYPE, item.EXPOSURE, item.SECURITYTYPEMAPPED, item.EXPOSUREMAPPED, item.OPENINGCOST, item.PURCHASES, item.SALES, item.AMORTIZATION, item.NETREALIZED, item.CLOSINGMARKETVALUE, item.EXCEPTIONS, item.RECONCILIATION]);

                    expensesTable.rows.add(null, newData);
                     if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                        currentWorksheet.getUsedRange().format.autofitColumns();
                        currentWorksheet.getUsedRange().format.autofitRows();
                     }
                   
                     var range = currentWorksheet.getRange("I2:K520");

                     range.dataValidation.rule = {
                         wholeNumber: {
                             formula1: 0,
                             operator: "GreaterThan"
                         }
                     };
                    
             
          await  context.sync();
            },
            (error) => {
                // A rejected prmise will execute this
                console.log('We have encountered an Error!'); // Log an error
            }
        );
       
      });
    } catch (error) {
      console.error(error);
    }
  }  
  
}
