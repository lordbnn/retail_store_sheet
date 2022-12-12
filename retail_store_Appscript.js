function onOpen() { 
    // insertRegion();
     customerInvoice();
     var ui = SpreadsheetApp.getUi();
     var ss= SpreadsheetApp.getActiveSpreadsheet();
     var debtorsSheet = ss.getSheetByName("Debtors Mastersheet");
     var debLr = debtorsSheet.getLastRow();
     var debLc = debtorsSheet.getLastColumn();
     
     var menuName = ui.createMenu("Company's Menu");
       var menuStock = ui.createMenu("Stocks");
       var menuCust = ui.createMenu("Customer");
       var menuWhouse = ui.createMenu("Warehouse");
       var menuReports = ui.createMenu("Reports");
       var menuHistory = ui.createMenu("History");
     Logger.log(debLr);
     
      menuStock.addItem("Add Stock", "addStockPrompt").addSeparator();
      menuStock.addItem("Drinks on Transit", "recordTransitDrinks").addSeparator();
      menuStock.addItem("Receive Stocks", "receiveStocks").addSeparator();
      menuStock.addItem("Change Price", "changePrice").addSeparator();
      menuStock.addItem("Record Damages", "recordDamages");
      menuName.addSubMenu(menuStock).addSeparator();
     
      menuCust.addItem("Add Customer", "addCustomer").addSeparator();
      menuCust.addItem("Customer Invoice", "customerInvoice").addSeparator();
      menuCust.addItem("Cash-In", "cashIn").addSeparator();
      menuCust.addItem("Pay Invoice", "invoiceByCustomer");
      menuName.addSubMenu(menuCust).addSeparator();
     
      menuWhouse.addItem("Expenditure", "expenditures");
      //menuReports.addItem("Dashboard", "#");
      menuName.addSubMenu(menuWhouse).addSeparator();
     
      menuReports.addItem("Stocks Report", "stocksReport").addSeparator();
      menuReports.addItem("Finance", "financeReport").addSeparator();
      menuReports.addItem("Annual Income", "incomeStatement");
      menuName.addSubMenu(menuReports).addSeparator();
     
      menuHistory.addItem("Invoice by Customer", "invoiceByCustomer").addSeparator();
      menuHistory.addItem("Invoice by date", "invoiceByDate").addSeparator();
      menuHistory.addItem("Invoice by Receipt No.", "invoiceByReceiptNo").addSeparator();
      menuHistory.addItem("Received Stocks", "receivedStocksHistory").addSeparator();
      menuHistory.addItem("Customer Payment History", "customerPaymentHistory").addSeparator();
      menuHistory.addItem("Drinks on Transit History", "transitHistory");
      menuName.addSubMenu(menuHistory);
      menuName.addToUi();
     
     ss.getSheetByName("Cash In").getRange(1, 4).setFormula("=Today()");
     ss.getSheetByName("Recieve Stock").getRange(1, 6).setFormula("=Today()");
     ss.getSheetByName("Recieve Stock").getRange(7, 6).setFormula('=ARRAYFORMULA(IFERROR(IF(C7:C="","",(C7:C)*(E7:E))))');
     ss.getSheetByName("Customer invoice").getRange(4, 8).setFormula("=Today()");
     ss.getSheetByName("Debtors Mastersheet").getRange(2, 1,debLr,1).setFormula("=ROW($A2)-ROW($A$1)");
     //ss.getSheetByName("Recieve Stock").getRange(8, 1,debLr,1).setFormula("=ROW($A4)-ROW($A$3)");
     
     
     var ssStock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
   
     var Lr = ssStock.getLastRow()-1;
     var Lc = ssStock.getLastColumn();
     //ss.getRange(2, 6,Lr,1).setFormula('=IF(RIGHT(A2,3)="PET","Pets",(IF(RIGHT(A2,3)="CAN","Cans","Bottles")))');
    
     
   
     invoiceFormulas();
     reportQuery();
     drinksOnTransitNewAmtFormula();
   }
   
   function hideSheets(activeSheet){
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     ss.getSheetByName(activeSheet).activate();
     var sheetsCount = ss.getNumSheets();
     var sheets = ss.getSheets();
     for(var i=0;i<sheetsCount;i++){
       var sheet = sheets[i];
       var fetchSheet =  sheet.getName()
       if (fetchSheet != activeSheet){
         sheet.hideSheet();
         //Logger.log(fetchSheet);
       }
      
     }
     
     
   }
   
   //VALIDATION FUNCTION
   function applyValidationToCell(list,cell){
     var rule = SpreadsheetApp
     .newDataValidation()
     .requireValueInList(list)
     .setAllowInvalid(false)
     .build();
     
     cell.setDataValidation(rule);
   
   }
   
   function expenditures(){ hideSheets("Expenditure");}
   function recordTransitDrinks(){ hideSheets("Record Transit Drinks");}
   function receiveStocks(){ hideSheets("Recieve Stock");}
   function changePrice(){ hideSheets("Change Price");}
   function customerInvoice(){ hideSheets("Customer invoice");}
   function cashIn(){ hideSheets("Cash In");}
   function recordDamages(){ hideSheets("Record Damages");}
   function stocksReport(){ hideSheets("Stocks Report");}
   function financeReport(){ hideSheets("Finance Report");}
   function incomeStatement(){ hideSheets("Income Statement");}
   function invoiceByCustomer(){ hideSheets("Invoice by Customer");}
   function invoiceByDate(){ hideSheets("Invoice by Date");}
   function invoiceByReceiptNo(){ hideSheets("Invoice by Receipt No.");}
   function receivedStocksHistory(){ hideSheets("Received Stocks History");}
   function customerPaymentHistory(){ hideSheets("Customer Finance History");}
   function transitHistory(){ hideSheets("Transit History");}
   
   function reportQuery(){
   
     var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ReportQuery");
     var ssRprt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stocks Report");
     var ssInvc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InvoiceWeeklyReport");
     var lrInv = ssInvc.getLastRow()-1;
     var lr = ss.getLastRow()-1;
     var lrSum = ss.getLastRow()-1;
     var lc = ss.getLastColumn();
     var lcRprt = ssRprt.getLastColumn();
     ss.getRange(2, 2, lr,1).setFormula("=SUMIF(stockWeeklyReport!$I$3:$I,A2,stockWeeklyReport!$J$3:$J)-SUMIF(InvoiceWeeklyReport!$AB$3:$AB,A2,InvoiceWeeklyReport!$AE$3:$AE)");
     ss.getRange(2, 3, lr,1).setFormula('=IFERROR(SUMIF(stockWeeklyReport!$B$2:$B,A2,stockWeeklyReport!$C$2:$C),0)');
     ss.getRange(2, 4, lr,1).setFormula('=SUMIF(InvoiceWeeklyReport!$C$2:$C,A2,InvoiceWeeklyReport!$F$2:$F)');
     ss.getRange(2, 5, lr,1).setFormula('=SUMIF(InvoiceWeeklyReport!$C$2:$C,A2,InvoiceWeeklyReport!$H$2:$H)');
     ss.getRange(2, 6, lr,1).setFormula('=B2+C2-D2');
     ss.getRange(2, 7, lr,1).setFormula('=IF(A2="","",VLOOKUP(A2,Stock!$A$2:$D,4,0))');
    
   
     //ss.getRange(lr+1,2,1,lc-1).setBackground('black').setFontColor('white').setFormula('=SUM(B2:'+'B'+lrSum+')');
    // ssRprt.getRange(lr,2,1,lc).setBackground('black').setFontColor('white');
   
        
          
          }
   
   
     
   
                            
   //CUSTOMER INVOICE DISCOUNT (%) & DISCOUNT INPUTS
   function onEdit(e) {
     var activeCell = e.range;
     var row = e.range.getRow();
     var col = e.range.getColumn();
     var cellVal = activeCell.getValue();
     var ss = e.source.getSheetByName("Customer invoice");
     var ssCash = e.source.getSheetByName("Cash In");
     // var ssStock = e.source.getSheetByName("Recieve Stock");
     var ssRprtQ = e.source.getSheetByName("ReportQuery");
     var ssRprt = e.source.getSheetByName("Stocks Report");
     var expRprt = e.source.getSheetByName("Expenditure");
     var breakages = e.source.getSheetByName("Record Damages");
     var invoiceByCust = e.source.getSheetByName("Invoice by Customer");
     var breakRow = breakages.getRange(row,2).getValue();
     var breakRowNum = breakages.getRange(row,3).getValue();
     // var lrRprt = ssRprt.getLastRow()-1;
     var lr = ssRprtQ.getLastRow()-1;
     var lc = ssRprtQ.getLastColumn();
    //  var getValues = ssRprtQ.getRange(2,1,lr,lc).getValues();
     var priceCol = ss.getRange(row,7).getValue();
     var qtyCol = ss.getRange(row,6).getValue();
     var getPercntg = ss.getRange(row,8).getValue();
     var getOrdnry = ss.getRange(row,9).getValue();
     var percntDiscnt = ss.getRange(row,8);
     var ordnryDisct = ss.getRange(row,9);
     var salesType = ss.getRange(row,5);
     var ssDropDownPage = e.source.getSheetByName("Invoice Mastersheet");
     var runOptions = ssDropDownPage.getRange("M2:N").getValues();
     // var qtyPurchased = ss.getRange(row,6); 
     //var regionCol = ss.getRange(row,13);
   
     
     
     var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
   
             if(col === 3 && row > 3){
   
             
          } 
     
     
     if(activeSheet==="Cash In"){
       
       if(col === 3 && row > 3){
   
          if (cellVal === ""){
               ssCash.getRange(row,4).clearContent()
              
             }else{
   
       
         ssCash.getRange(row,4).setFormula("INDEX('Customer List'!$B$3:$B,MATCH("+"C"+row+",'Customer List'!$E$3:$E,0))");
        ssCash.getRange(row,6).setFormula("SUMIF('Invoice Mastersheet'!$N$2:$N,"+"E"+row+",'Invoice Mastersheet'!J2:J)-SUMIF('Cash-In Mastersheet'!$E$2:$E,"+"E"+row+",'Cash-In Mastersheet'!G2:G)")
         
         
         }
        
   
              if (cellVal === ""){
               ssCash.getRange(row,5).clearContent().clearDataValidations();
              
             }else{
             ssCash.getRange(row,5).clearContent()
             var filteredOptions = runOptions.filter(function(o){return o[0] === cellVal});
             var listToApply = filteredOptions.map(function(o){return o[1]});
             var cell = ssCash.getRange(row,5);
             
               applyValidationToCell(listToApply,cell);
               Logger.log(listToApply);
               console.log(runOptions);
             }
   
             } 
        if(col === 1 && row > 3){
         ssCash.getRange(row,2).clearContent();
         ssCash.getRange(row,3).clearContent();
         ssCash.getRange(row,4).clearContent();
         ssCash.getRange(row,5).clearContent();
             } 
       
   
   
   
      
     }
     
     if(activeSheet=="Stocks Report"){
       
       if(col === 4 && row ===1){
         ssRprtQ.getRange(2, 2, lr, lc).clear();
         
         ssRprt.getRange(4,2,lr+4,lc).setBackground(null);
         ssRprt.getRange(4,2,lr+4,lc).clear();
         reportQuery();
         ssRprt.getRange(4,2,1,1).setFormula('=QUERY(ReportQuery!A2:L)');    
   
             } 
       
        if(col === 6 && row ===2){
         ssRprtQ.getRange(2, 2, lr, lc).clear();
         
         ssRprt.getRange(4,2,lr+4,lc).setBackground(null);
         ssRprt.getRange(4,2,lr+4,lc).clear();
         reportQuery();
   
        ssRprt.getRange(4,2,1,1).setFormula('=QUERY(ReportQuery!A2:L)');      
    
             }
        if(col === 9 && row ===2){
         ssRprtQ.getRange(2, 2, lr, lc).clear();
         
         ssRprt.getRange(4,2,lr+4,lc).setBackground(null);
         ssRprt.getRange(4,2,lr+4,lc).clear();
         reportQuery();
        // Logger.log(lr);
         //Logger.log(lrRprt);
         ssRprt.getRange(4,2,1,1).setFormula('=QUERY(ReportQuery!A2:L)');      
       
             }
     }
   
     
     
        if(activeSheet==="Change Price"){
              if(col === 3 && row > 7){
                priceChange(); } 
             } 
     
     
      if(activeSheet==="Record Damages"){
       var hack = '"'+breakRow+'"';//breakRowNum
       if(col === 3 && row > 3){
         breakages.getRange(row,4).setFormula("="+breakRowNum+"*"+"VLOOKUP("+hack+",'COGS Mastersheet'!$H$2:$I,2,0)");
         breakages.getRange(row,5).setFormula('$B$1');
             } 
        if(col === 2 && row > 3){
     breakages.getRange(row,3).clearContent();
     breakages.getRange(row,4).clearContent();
     breakages.getRange(row,5).clearContent();
      }
      } 
   
     if(activeSheet==="Expenditure"){
       
       if(col === 2 && row > 3){
         expRprt.getRange(row,6).setFormula('$C$2');
         
             } 
        if(col === 1 && row > 3){
         expRprt.getRange(row,2).clearContent();
         expRprt.getRange(row,3).clearContent();
         expRprt.getRange(row,4).clearContent();
         expRprt.getRange(row,5).clearContent();
         expRprt.getRange(row,6).clearContent();
        
             } 
      
     }
     
     
       if(activeSheet==="Invoice by Customer"){
         var invByCustLr = invoiceByCust.getLastRow()-6; 
              
         if(col===6 && row > 6){invoiceByCust.getRange(row,8).setFormula('$B$3');}
         invoiceByCust.getRange('$F$7').setFormula('ARRAYFORMULA(IF(E7:E="","",(C7:C+D7:D)))')
         invoiceByCust.getRange('$G$7').setFormula('ARRAYFORMULA(IF(A7:A="","",(A7:A)+30))')
         
   
      
       }
     
      if(activeSheet=="Customer invoice"){
   
          if(col === 3 && row > 12){
       onEditInvoiceFormula();
       salesType.setValue("Wholesale");
       
             } 
        
     if(col === 8 && row > 12){
       ordnryDisct.setFormula((getPercntg/100)*priceCol*qtyCol);
             } 
     
      if(col === 9 && row > 12){
       percntDiscnt.setFormula((getOrdnry/priceCol)*100);
          
            }
    /* 
      if(col === 3 && row > 12){
        ordnryDisct.clearContent();
        percntDiscnt.clearContent();
        qtyPurchased.clearContent();
        regionCol.clearContent();
             }*/
   
       
     if(col === 5 && row > 12){
       
       ss.getRange(row,13).setFormula('=IF(C'+row+'="","",$D$3)');
       ss.getRange(row,14).setFormula('=IF(C'+row+'="","",$D$8)');
       ss.getRange(13,17,1,1).setFormula('=IF(C'+row+'="","",$D$8)');
       ss.getRange(13,15,1,1).setFormula('=IF(C'+row+'="","",$F$9)');
       ss.getRange(13,16,1,1).setFormula('=IF(C'+row+'="","",$D$4)');
       //ss.getRange()
       
             } 
     if(row > 12) {
   
     if(ss.getRange(row,4).getValue() < ss.getRange(row,6).getValue()){
         ss.getRange(row,6).setBackgroundRGB(255, 0, 0);
     }else {ss.getRange(row,6).setBackground(null);}
       
     }
      }
       
   }
   
   
   //CUSTOMER LIST FORMULA
   
   function custListFormula(){
       var ssCustomerList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer List");
       var customerLr = ssCustomerList.getLastRow()-2;
       ssCustomerList.getRange(3, 4,customerLr,1).setFormula("=SUMIF('Cash-In Mastersheet'!$C$2:$C,E3,'Cash-In Mastersheet'!$G$2:$G)-SUMIF('Invoice Mastersheet'!$M$2:M,E3,'Invoice Mastersheet'!$J$2:$J)")
       ssCustomerList.getRange(3, 1,customerLr,1).setFormula("=ROW(A3)-ROW($A$2)");
       ssCustomerList.getRange(3, 5,customerLr,1).setFormula('=IF(B3="","","TOFF00"&row()-2)');
   
   }
   
   
   //CASH-IN & EXPENDITURE TRANSFER TEMPLATE
   
   function sheetTransfer(fetch,master,ROWcount, STARTrow){
     //var rowCount = 1;
     var startRow = STARTrow;
     var ssfetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(fetch);
     var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(master);
     var fetchLr = ssfetch.getLastRow()-ROWcount;
     var masterLr = ssMaster.getLastRow();
    // Logger.log(masterLr);
     var fetchLc = ssfetch.getLastColumn();
     var masterLc = ssMaster.getLastColumn();
     var fetchContents = ssfetch.getRange(startRow, 1, fetchLr, masterLc);
     var fetchValues = fetchContents.getValues();
     var sheetFormulas = fetchContents.getFormulas();
     
     if (masterLr == 1){ssMaster.getRange(2, 1, fetchLr, masterLc).setValues(fetchValues);}else{
       ssMaster.getRange(masterLr+1, 1, fetchLr, masterLc).setValues(fetchValues);}
           
     
       fetchContents.clearContent();
       fetchContents.setFormulas(sheetFormulas);
       
   
   }
   
   function cashinTransfer(){ sheetTransfer("Cash In","Cash-In Mastersheet",1,4);custListFormula();}
   function breakagesTransfer(){ sheetTransfer("Record Damages","Breakages Mastersheet",1,4);
                                var ssBreakagesMS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Breakages Mastersheet");  
                                var ssBreakages = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Record Damages"); 
                                var lr = ssBreakagesMS.getLastRow()-1;
                                ssBreakages.getRange(4, 4, 10,5).clearContent();
                                ssBreakagesMS.getRange(2, 6, lr,1).setFormula('IF(A2="","",YEAR(A2))');  
                                ssBreakagesMS.getRange(2, 7, lr,1).setFormula('IF(A2="","",TEXT(A2,"MMMM"))'); 
   }
   
   function expenseTransfer(){ sheetTransfer("Expenditure","Expenditure Mastersheet",1,4);invoiceFormulas();}
   function transitDrinksTransfer(){ sheetTransfer("Record Transit Drinks","DrinksOnTransitMS",5,6);drinksOnTransitNewAmtFormula();}
   
   
   function customerInvoiceBtn() {
   
     var ssInvoiceFetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer invoice");
     var ssInvoiceMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Mastersheet");
     var invoiceLr = ssInvoiceFetch.getLastRow()-11;
     var masterLr = ssInvoiceMaster.getLastRow();
    // Logger.log(masterLr);
     var invoiceLc = ssInvoiceFetch.getLastColumn();
     var masterLc = ssInvoiceMaster.getLastColumn();
     var masterValues = ssInvoiceMaster.getRange(2, 1, invoiceLr, masterLc).getValues();
     var invoiceContents = ssInvoiceFetch.getRange(13, 1, invoiceLr, masterLc);
     var invoiceValues = invoiceContents.getValues();
     var sheetFormulas = invoiceContents.getFormulas();
    
     
     if (masterLr == 1){ssInvoiceMaster.getRange(2, 1, masterLr,masterLc).setValues(invoiceValues);}else{
       ssInvoiceMaster.getRange(masterLr+1, 1, invoiceLr,masterLc).setValues(invoiceValues);
           
     
       invoiceContents.clearContent();
       ssInvoiceFetch.getRange(4, 4).clearContent();
       invoiceContents.setFormulas(sheetFormulas);
   
   }
     invoiceFormulas();
    
   
   }
   
   
   function onEditInvoiceFormula(){
       var ssInvoiceFetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer invoice");
     var ssInvoiceMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Mastersheet");
     var ssInvoiceCOGS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COGS Mastersheet");
     var ssExpendQuery = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenditure Mastersheet");
     var lrExpend = ssExpendQuery.getLastRow()-1;
     var COGSlr = ssInvoiceCOGS.getLastRow()-1;
     var invoiceLr = ssInvoiceFetch.getLastRow()-12;
     var masterLr = ssInvoiceMaster.getLastRow()-1;
    // Logger.log(masterLr);
     var invoiceLc = ssInvoiceFetch.getLastColumn();
     var masterLc = ssInvoiceMaster.getLastColumn();
   
     ssInvoiceFetch.getRange(13, 1, invoiceLr, 1).setFormula('=IF(J13="","",$F$9)');
     ssInvoiceFetch.getRange(13, 2, invoiceLr, 1).setFormula('=IF(J13="","",$D$4)');
     ssInvoiceFetch.getRange(13, 4, invoiceLr, 1).setFormula('=IFERROR(VLOOKUP(C13,Stock!$A$2:$B,2,0),"")');
     ssInvoiceFetch.getRange(13, 7, invoiceLr, 1).setFormula("=IFERROR(VLOOKUP(C13,'Price Mastersheet'!$B$2:$H,7,0))");
     ssInvoiceFetch.getRange(13, 7, invoiceLr, 1).setFormula("=IFERROR(INDEX('Price Mastersheet'!$F$2:F,SUMPRODUCT(MAX(ROW('Price Mastersheet'!$B$2:C)*(C13='Price Mastersheet'!$B$2:B)*(E13='Price Mastersheet'!$C$2:C))-1)))");
     ssInvoiceFetch.getRange("J13").setFormula('=ARRAYFORMULA(IF(F13:F="","",((G13:G)*(F13:F))-((I13:I)*(F13:F))))');
     ssInvoiceFetch.getRange(13, 11, 1, 1).setFormula('=$D$9');
     ssInvoiceFetch.getRange(13, 12, 1, 1).setFormula('=$D$5');
   }
   
   
   
   function invoiceFormulas(){
   
     var ssInvoiceFetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer invoice");
     var ssInvoiceMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Mastersheet");
     var ssInvoiceCOGS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COGS Mastersheet");
     var ssExpendQuery = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenditure Mastersheet");
     var lrExpend = ssExpendQuery.getLastRow()-1;
     var COGSlr = ssInvoiceCOGS.getLastRow()-1;
     var invoiceLr = ssInvoiceFetch.getLastRow()-11;
     var masterLr = ssInvoiceMaster.getLastRow()-1;
    // Logger.log(masterLr);
     var invoiceLc = ssInvoiceFetch.getLastColumn();
     var masterLc = ssInvoiceMaster.getLastColumn();
   
    // ssInvoiceFetch.getRange(13, 8, invoiceLr-11, 2).clearContent();
     ssInvoiceFetch.getRange(13, 1, invoiceLr, masterLc).setBackground(null);
     ssInvoiceFetch.getRange(5, 4).setFormula("=IFERROR(VLOOKUP(D4,'Customer List'!B3:D,3,0),0)");
     ssInvoiceFetch.getRange(9, 4).setFormula("=SUM(J13:J)");
   
     ssInvoiceCOGS.getRange(2, 10, COGSlr, 1).setFormula('=F2*(G2-VLOOKUP(C2,$H$2:$I,2,0))');
     ssInvoiceCOGS.getRange(2, 13, COGSlr, 1).setFormula("=L2+K2+SUMIFS('Cash-In Mastersheet'!$G$2:$G,'Cash-In Mastersheet'!$C$2:$C,B2,'Cash-In Mastersheet'!$A$2:$A,A2)");
     ssInvoiceCOGS.getRange(2, 14, COGSlr, 1).setFormula('=if(M2<0,M2,0)')
   
    // ssInvoiceFetch.getRange(13, 13,invoiceLr,1).setFormula('=$G$9');
     ssInvoiceFetch.getRange(8, 4).setFormula('=IF(J13 = "","",right(year(H4), 2) & text(month(H4), "00") & text(day(H4), "00") & left(iferror(regexextract(D13 & F13 & J13, "\d+")) &'+masterLr+'+2'+', 6))');
     ssInvoiceMaster.getRange(2, 1, invoiceLr, masterLc).clearFormat();
     //ssInvoiceMaster.getRange("R2").setFormula('ARRAYFORMULA(IFERROR(lookup(A2:A,wkcode!B$3:$C,wkcode!$A$3:A)))');
     //ssInvoiceMaster.getRange("S2").setFormula('=ARRAYFORMULA(IFERROR(IF(E2:E="Loan",(VLOOKUP($C2:C,Stock!$A$2:$G,7,0)*$F2:F),(VLOOKUP($C2:C,Stock!$A$2:$G,6,0)*$F2:F))))');
     ssInvoiceCOGS.getRange(2, 16, COGSlr, 1).setFormula('IF(A2="","",YEAR(A2))');
     ssInvoiceCOGS.getRange(2, 17, COGSlr, 1).setFormula('IF(A2="","",TEXT(A2,"MMMM"))');
     ssInvoiceCOGS.getRange(2, 18, COGSlr, 1).setFormula('=IF(Stock!F2="Bottles",Vlookup(Stock!A2,Stock!$A$2:$A,1,0),"")');
     ssExpendQuery.getRange(2, 7, lrExpend, 1).setFormula('IF(A2="","",YEAR(A2))');
     ssExpendQuery.getRange(2, 8, lrExpend, 1).setFormula('IF(A2="","",TEXT(A2,"MMMM"))');
   
   
   }
   
   function queryFunction(){
     var ssDrop = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("invoiceQueries");
     var ssHistory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CustomerFinHistoryMS");
     var ssFetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice by Customer");
     var lr = ssDrop.getLastRow()-1;
     var Histlr = ssHistory.getLastRow()-1;
     var fetchlr = ssFetch.getLastRow()-6;
     
     var statusRange = ssDrop.getRange(2, 20,lr,1);
     var DifferenceRange = ssDrop.getRange(2, 19,lr,1);
     var fetchName = ssFetch.getRange(3, 2,1,1).getValue();
     
    // ssHistory.getRange(Histlr+2, 7,fetchlr,1).setValue(fetchName);
     statusRange.setFormula('=IF(S2>0,"EXCESS",if(S2="","",IF(S2=0,"PAID","UNPAID")))');
     DifferenceRange.setFormula('=U2-K2');
     Logger.log(lr);
   }
   
   
   
   function invoiceByCustomerProcessor(){
       var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("invoiceQueries");
       var ssSlave = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice by Customer");
       var masterLr = ssMaster.getLastRow()-1;
       var slaveLr = ssSlave.getLastRow();
     
       var custNameSlave = ssSlave.getRange("B3").getValue();
       var custNameMaster = ssMaster.getRange(2, 2, masterLr,1).getValues();
     
       var invNumberSlave = ssSlave.getRange(7, 2,slaveLr,1).getValues();
       var invNumberMaster = ssMaster.getRange(2, 17, masterLr,1).getValues();
     
       var paidAmountMaster = ssMaster.getRange(2, 21,masterLr,1);
       var paidAmountMasterGetVal = ssMaster.getRange(2, 21,masterLr,1).getValues();
   
    
   
     for(var i=0; i<slaveLr; i++){
         
       for(var j=0; j<masterLr; j++){
        
           
         if(custNameSlave===custNameMaster[j][0] && invNumberSlave[i][0].toString()===invNumberMaster[j][0].toString()){//if both customer names and invoice numbers are thesame...
           
           
           var paidAmountSlave = ssSlave.getRange(i+7, 6).getValue(); //get the values of 'paid amount' from the 'invoice by customer sheet'
          
           
          var getAmountRangemaster = ssMaster.getRange(j+2, 21).getValue(); //Add it to the 'paid amount' of that row in the invoiceQueries sheet
           ssMaster.getRange(j+2, 21).setFormula(getAmountRangemaster+paidAmountSlave);
   
                              
                              } 
          
         
   }}
   
   queryFunction();
   createCustomerHistory();
   }
   
   
   
   
   function createCustomerHistory(){
     var fetchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice by Customer');
     var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerFinHistoryMS');
     var lr = fetchSheet.getLastRow()-6;
     var lc = fetchSheet.getLastColumn();
     var copyRange = fetchSheet.getRange(7, 1,lr,lc).getValues();
     for(var i=0; i<lr; i++){
       //var copyRange = fetchSheet.getRange(i, 1,1,6).getValues();
             masterSheet.appendRow(copyRange[i]);}
   
   
   }
   
   
   
   function receiveStockBtn() {
   
     var ssReceive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recieve Stock");
     var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock Mastersheet");
     var receiveLr = ssReceive.getLastRow()-6;
     var masterLr = ssMaster.getLastRow();
    // Logger.log(masterLr);
     var receiveLc = ssReceive.getLastColumn();
     var masterLc = ssMaster.getLastColumn();
     var masterValues = ssMaster.getRange(2, 1, receiveLr, masterLc).getValues();
     var receiveContents = ssReceive.getRange(7, 1, receiveLr, masterLc);
     var receiveValues = receiveContents.getValues();
     
    
     if (masterLr == 1){
       ssMaster.getRange(2, 1, receiveLr, masterLc).setValues(receiveValues);}else{
       ssMaster.getRange(masterLr+1, 1, receiveLr, masterLc).setValues(receiveValues);
       
       }
     
     receiveContents.clearContent();
     drinksOnTransitNewAmtFormula();
     ssReceive.getRange(1, 5, 1, 2).setFormula("=Today()");
     ssReceive.getRange(7, 6).setFormula('=ARRAYFORMULA(IFERROR(IF(C7:C="","",(C7:C)*(E7:E))))')
     ssMaster.getRange(2, 1, receiveLr, masterLc).clearFormat();
    
     
   }
   
   
   //CHANGE PRICE SHEET
   function saveNewPriceBtn(){
     
     var ssPrice = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Change Price");
     var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Price Mastersheet");
     var masterLr = ssMaster.getLastRow();
     var priceLr = ssPrice.getLastRow()-7;
     //Logger.log(priceLr);
     var masterLc = ssMaster.getLastColumn();
     var priceLc = ssPrice.getLastColumn();
     var masterValues = ssMaster.getRange(2, 1, priceLr, masterLc).getValues();
     var priceContents = ssPrice.getRange(8, 1, priceLr, masterLc);
     var priceValues = priceContents.getValues();
    // var sheetFormulas = priceContents.getFormulas();
     //Logger.log(masterLr);
     
   
      
      if (masterLr == 1){
       ssMaster.getRange(2, 1, priceLr, masterLc).setValues(priceValues);}else{
       ssMaster.getRange(masterLr+1, 1, priceLr, masterLc).setValues(priceValues);
       
       }
     
     
     priceContents.clearContent();
     ssMaster.getRange(2, 8, masterLr, 1).setFormula('=INDEX($F$2:F,SUMPRODUCT(MAX(ROW($B$2:B)*(B2=$B$2:B))-1))');
   }
   
   
   function priceChange(){
     var ssPrice = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Change Price");
     var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Price Mastersheet");
     var masterLr = ssMaster.getLastRow()-1;
     var priceLr = ssPrice.getLastRow()-7;
     
     //ssMaster.getRange(2, 1, priceLr, masterLc).clearFormat();
     ssPrice.getRange(8, 1, priceLr, 1).setFormula('=IF(F8="","",$E$1)');
     ssPrice.getRange(8, 4, priceLr, 1).setFormula("=IFERROR(INDEX('Price Mastersheet'!$F$2:F,SUMPRODUCT(MAX(ROW('Price Mastersheet'!$B$2:C)*(B8='Price Mastersheet'!$B$2:B)*(C8='Price Mastersheet'!$C$2:C))-1)))");
     ssPrice.getRange(8, 5, priceLr, 1).setFormula('=IFERROR(VLOOKUP(B8,Stock!$A$2:$B,2,0),"")');
     ssPrice.getRange(8, 7, priceLr, 1).setFormula('=IFERROR(IF(F8="","",ROUND(((F8-D8)/D8),2)),"")');
     
   }
   
   
   //FORMULA TO DETECT BRANDS MATERIALS IN ORDER TO QUERY THE BOTTLES TO FETCH EMPTIES CALCULATION FROM
   function stockFormula(){
     var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
     var Lr = ss.getLastRow()-1;
     var Lc = ss.getLastColumn();
     var stringWord = "'Invoice Mastersheet'";
     var stringWord2 = "'Stock Mastersheet'";
      ss.getRange("G2").setFormula('=ARRAYFORMULA(IF(D2:D="","",$D$2:$D-$E$2:$E))'); //AvgProfitLoan
   
     ss.getRange("F2").setFormula('=ARRAYFORMULA(IF(C2:C="","",$C$2:$C-$E$2:$E))'); //AvgProfitCash
   
     ss.getRange(2, 5,Lr,1).setFormula('=IF(A2="","",SUMIF('+stringWord2+'!$B$2:$B,$A2,'+stringWord2+'!$F$2:$F)/SUMIF('+stringWord2+'!$B$2:$B,$A2,'+stringWord2+'!$C$2:$C))'); //AvgCost
   
     ss.getRange(2, 4,Lr,1).setFormula("=IFERROR(INDEX('Price Mastersheet'!$F$2:F,SUMPRODUCT(MAX(ROW('Price Mastersheet'!$B$2:C)*(A2='Price Mastersheet'!$B$2:B)*($D$1='Price Mastersheet'!$C$2:C))-1)))"); //Loan column
   
     ss.getRange(2, 3,Lr,1).setFormula("=IFERROR(INDEX('Price Mastersheet'!$F$2:F,SUMPRODUCT(MAX(ROW('Price Mastersheet'!$B$2:C)*(A2='Price Mastersheet'!$B$2:B)*($C$1='Price Mastersheet'!$C$2:C))-1)))"); // Cash column
   
     ss.getRange(2, 2,Lr,1).setFormula("=(SUMIF('Stock Mastersheet'!$B$2:$B,A2,'Stock Mastersheet'!$C$2:$C))-(SUMIF('Invoice Mastersheet'!$C$2:$C,A2,'Invoice Mastersheet'!$F$2:$F))-SUMIF('Breakages Mastersheet'!$B$2:$B,A2,'Breakages Mastersheet'!$C$2:$C)");
     //Logger.log(Lr)
     
   }
   
   
   
   
   function drinksOnTransitNewAmtFormula(){
     var ssDrop = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DrinksOnTransitMS");
     var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock Mastersheet");
     var dropLr = ssDrop.getLastRow()-5;
     var dropLc = ssDrop.getLastColumn();
     var masterLr = ssMaster.getLastRow()-1;
     var masterLc = ssMaster.getLastColumn();
     var masterBrands = ssMaster.getRange(2,2,masterLr,1).getValues();
     var dropBrands = ssDrop.getRange(6,2,dropLr,1).getValues();
     var shipmntMaster = ssMaster.getRange(2,4,masterLr,1).getValues();
     var shipmntDrop = ssDrop.getRange(6,4,dropLr,1).getValues();
     var amountMaster = ssMaster.getRange(2,6,masterLr,1).getValues();
     var amountDrop = ssDrop.getRange(6,5,dropLr,1).getValues();
     var qtyMaster = ssMaster.getRange(2,3,masterLr,1).getValues();
     var qtyDrop = ssDrop.getRange(6,3,dropLr,1).getValues();
   
     //var Lrv = ss.getRange(2, 3,Lr,1).getValues().length;
     
     for(i=0; i<dropLr; i++){
         var newAmtDrop = ssDrop.getRange(i+6,6);
         var newQtyDrop = ssDrop.getRange(i+6,7);
       
       for(j=0; j<masterLr; j++){
   
         if(dropBrands[i][0]===masterBrands[j][0] && shipmntDrop[i][0]===shipmntMaster[j][0]){
          
           
           var amountDifference = amountMaster[j][0]-amountDrop[i][0];
           var qtyDifference = qtyMaster[j][0]-qtyDrop[i][0];
           
           newAmtDrop.setValue(amountDifference);
           newQtyDrop.setValue(qtyDifference);
           
   
           
         
         }
       
       }
   
         }
   
   drinksOnTransitQueryFormula();
   }
   
   
   function drinksOnTransitQueryFormula(){
   
     var ssQuerySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("drinksOnTransitQuery");
     var lr = ssQuerySheet.getLastRow()-1;
     ssQuerySheet.getRange(2, 8, lr, 1).setFormula('=IF(G2="","",IF(F2=0,IF(G2=0,"CLEARED","UNCLEARED"),"UNCLEARED"))');
   }
   
   
   
   //PROMPTS (ADD STOCKS & ADD CUSTOMER)
   
   function addStockPrompt(){
     var ss = SpreadsheetApp.getUi();
     var input = ss.prompt("ADD STOCK", "Please input the name of the brand you'd like to add", ss.ButtonSet.OK_CANCEL);
     var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
     var stockLr = stockSheet.getLastRow()+1;
     Logger.log(stockLr);
     
     if(input.getSelectedButton()==ss.Button.OK){
      
       var promptText = input.getResponseText().toUpperCase();
         
       stockSheet.getRange(stockLr, 1).setValue(promptText);
       //changePrice();
       receiveStocks();
       stockFormula();
          
        
     
     }else if(input.getSelectedButton()==ss.Button.CANCEL){
   //don nothing
     
     }else if(input.getSelectedButton()==ss.Button.CLOSE){
     //do nothing
     }
   }
   
   function addCustomer(){
     var ss = SpreadsheetApp.getUi();
     var input = ss.prompt("ADD CUSTOMER", "Please input the NEW customer's name", ss.ButtonSet.OK_CANCEL);
     var custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer List");
     var custLr = custSheet.getLastRow()+1;
     //Logger.log(custLr);
     
     if(input.getSelectedButton()==ss.Button.OK){
       var promptText = input.getResponseText().toUpperCase();
       
       custSheet.getRange(custLr, 2).setValue(promptText);
       customerInvoice()
       custListFormula();  
     
     
     }else if(input.getSelectedButton()==ss.Button.CANCEL){
     //do nothing
     
     }else if(input.getSelectedButton()==ss.Button.CLOSE){
     //do nothing
     
     }
   }
   
   
   function insertRegion(){
     var ss = SpreadsheetApp.getUi();
     var input = ss.prompt("INSERT BRANCH", "What branch are you?", ss.ButtonSet.OK_CANCEL);
     var custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customer invoice");
    
     
     if(input.getSelectedButton()==ss.Button.OK){
       var promptText = input.getResponseText().toUpperCase();
     //  if(promptText!="OBUDU" || promptText!="BEKWARA" || promptText!="obudu" || promptText!="bekwara"){insertRegion();}else{
       custSheet.getRange("G9").setValue(promptText);
         customerInvoice(); // }
   
     
     
     }else if(input.getSelectedButton()==ss.Button.CANCEL){
     insertRegion();
     
     }else if(input.getSelectedButton()==ss.Button.CLOSE){
     insertRegion();
     
     }
   }
   
   
   