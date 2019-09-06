import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'ng-xlsx-up';
  data: AOA = [ [1, 2], [3, 4] ];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName = 'SheetJS.xlsx';
  headers = [];
  selectedHeader = '';

  onFileChange(evt: any) {
    const target: DataTransfer = (evt.target) as DataTransfer;
    if (target.files.length !== 1) {
      throw new Error('Connot use multiple files');
    }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      this.data = (XLSX.utils.sheet_to_json(ws, {header: 1})) as AOA;
    };
    reader.readAsBinaryString(target.files[0]);
  }

  readToConsole() {
    this.data[0].forEach((val, i) => {
      this.headers.push(val);
    });
    console.log(this.data[0]);
    console.log(this.headers);
  }

  export(): void {
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, this.fileName);
  }

  onSelectHeader(i: any) {
    this.data.forEach(row => {
      console.log(row[i]);
    });
  }
}
