import { Component } from '@angular/core';
import '@mescius/spread-sheets-angular';
import * as GC from "@mescius/spread-sheets";
import '@mescius/spread-sheets-io';
import '@mescius/spread-sheets-charts';
import '@mescius/spread-sheets-shapes';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})

export class AppComponent {
  title = "SJS-Angular-IO-Excel";
  hostStyle = {
    width: '95%',
    height: '600px'
  };
  private spread;
  private currentCustomerIndex = 0;
  columnWidth = 100;
  selectedFile: File | any = null;
  constructor() {
    this.spread = new GC.Spread.Sheets.Workbook();
  }

  // Initialize SpreadJS
  workbookInit($event: any) {
    this.spread = $event.spread;
  }
  
  // Get the selected Excel file
  // Get the selected Excel file
  selectedFileChange(e: any) {
    this.selectedFile = e.target.files[0];
}
  
  // Imported the Excel file into SpreadJS
  open() {
    var file = this.selectedFile;
        if (!file) {
            return;
        }
    // Specify the file type to ensure proper import
    const options: GC.Spread.Sheets.ImportOptions = {
      fileType: GC.Spread.Sheets.FileType.excel
    };

    this.spread.import(file, () => {
      console.log('Import successful');
    }, (e: any) => {
      console.error('Error during import:', e);
    }, options);
  }
  
  // Modify data imported
  // Modify data imported
  addCustomer() {
    // create new row and copy styles
    var newRowIndex = 34;
    var sheet = this.spread.getActiveSheet();
    sheet.addRows(newRowIndex, 1);  
    sheet.copyTo(32, 1, newRowIndex, 1, 1, 11, GC.Spread.Sheets.CopyToOptions.style);  
    // Define sample customer data
    var customerDataArrays = [
        ["Jessica Moth", 5000, 2000, 3000, 1300, 999, 100],
        ["John Doe", 6000, 2500, 3500, 1400, 1000, 20],
        ["Alice Smith", 7000, 3000, 4000, 1500, 1100, 0]
    ];
    // Get the current customer data array
    var currentCustomerData = customerDataArrays[this.currentCustomerIndex];
    // Add new data to the new row
    sheet.setArray(newRowIndex, 5, [currentCustomerData]);
    newRowIndex++;
    // Increment the index for the next button click
    this.currentCustomerIndex = (this.currentCustomerIndex + 1) % customerDataArrays.length;
  }
  // Save the instance as an Excel File
   // Save the instance as an Excel File
   save() {
    var fileName = 'Excel_Export.xlsx'
    this.spread.export(function (blob:any) {
      // save blob to a file
      saveAs(blob, fileName);
    }, function (e:any) {
        console.log(e);
    }, {
        fileType: GC.Spread.Sheets.FileType.excel
    });
  }
}