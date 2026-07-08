import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';


@Component({
  selector: 'app-home',
  imports: [CommonModule],
  templateUrl: './home.html',
  styleUrl: './home.css',
})
export class homeComponent {

  excel_data: any;
  constructor() {



  }
  ReadExcel(event: any) {
    let file = event.target.files[0];

    let fileReader = new FileReader();
    fileReader.readAsArrayBuffer(file);



    fileReader.onload = (e) => {

      var workbook = XLSX.read(fileReader.result, { type: 'array' })
      var sheetnames = workbook.SheetNames;
      this.excel_data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetnames[0]])
      console.log(this.excel_data)




    }



  }


}

