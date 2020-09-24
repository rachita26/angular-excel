import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'angular-excel';
  fileName = 'excelImport.xlsx';
  book: XLSX.WorkBook;
  sheet: XLSX.WorkSheet;

  people = [{
    'name': 'Noddy', 'age': '10', 'city': 'JSR', 'phone': '123'
  },
  {
    'name': 'Oswald', 'age': '15', 'city': 'BLR', 'phone': '456'
  },
  {
    'name': 'Bob', 'age': '30', 'city': 'KOL', 'phone': '789'
  },
  {
    'name': 'Pingu', 'age': '5', 'city': 'JSR', 'phone': '987'
  }];

  exportToExcel() {
    var table = document.getElementById('exportTable'); //to fetch the table element by id
    this.book = XLSX.utils.book_new(); //create a new book
    this.sheet = XLSX.utils.table_to_sheet(table); //convert the table to a sheet
    XLSX.utils.book_append_sheet(this.book, this.sheet, 'Sheet1'); //append sheet to book, with desired name

    //this.book = XLSX.utils.table_to_book(table); => can be used directly if sheet name specification not required (skip lines 31, 32 if used)

    XLSX.writeFile(this.book, this.fileName); //write the file with specified name
  }

}
