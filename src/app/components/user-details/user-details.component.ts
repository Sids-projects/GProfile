import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, FormArray } from '@angular/forms';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-user-details',
  templateUrl: './user-details.component.html',
  styleUrls: ['./user-details.component.scss'],
})
export class UserDetailsComponent implements OnInit {
  dataForm!: FormGroup;

  constructor(private formBuilder: FormBuilder) {}

  ngOnInit() {
    this.dataForm = this.formBuilder.group({
      formId: [''],
      location: [''],
      hotelDetails: this.formBuilder.array([this.createHotelDetail()]),
      pax: [''],
      adultCount: [''],
      childCount: [''],
      childage: [''],
      noOfRooms: [''],
      occupancy: [''],
      mealPlan: [''],
      cabName: [''],
      cabContact: [''],
      customerDetails: this.formBuilder.array([this.createCustomerDetail()]), // Initialize with one entry
    });
  }

  get hotelDetails(): FormArray {
    return this.dataForm.get('hotelDetails') as FormArray;
  }

  createHotelDetail(): FormGroup {
    return this.formBuilder.group({
      hotelName: [''],
      hotelContact: [''],
      advanceAmt: [''],
      checkIn: [''],
      checkOut: [''],
    });
  }

  addHotelDetail(): void {
    this.hotelDetails.push(this.createHotelDetail());
  }

  get customerDetails(): FormArray {
    return this.dataForm.get('customerDetails') as FormArray;
  }

  createCustomerDetail(): FormGroup {
    return this.formBuilder.group({
      cusName: [''],
      cusContact: [''],
    });
  }

  addCustomerDetail(): void {
    this.customerDetails.push(this.createCustomerDetail());
  }

  saveExcel() {
    const formData = this.dataForm.value;
    const rows: any[][] = [];

    Object.keys(formData).forEach((key) => {
      if (key !== 'hotelDetails' && key !== 'customerDetails') {
        rows.push([key, formData[key]]);
      }
    });

    rows.push(['hotelDetails']);
    formData.hotelDetails.forEach((detail: any) => {
      const hotelDetailRow: any[] = [];
      Object.keys(detail).forEach((key) => {
        hotelDetailRow.push(`${key}: ${detail[key]}`);
      });
      rows.push(hotelDetailRow);
    });

    rows.push(['customerDetails']);
    formData.customerDetails.forEach((detail: any) => {
      const customerDetailRow: any[] = [];
      Object.keys(detail).forEach((key) => {
        customerDetailRow.push(`${key}: ${detail[key]}`);
      });
      rows.push(customerDetailRow);
    });

    const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data'],
    };

    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    this.saveAsExcelFile(excelBuffer, 'formData');
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    });

    const excelUrl: string = window.URL.createObjectURL(data);

    const link: HTMLAnchorElement = document.createElement('a');
    link.href = excelUrl;
    link.download = `${fileName}.xlsx`;
    link.click();

    window.URL.revokeObjectURL(excelUrl);
  }
}
