package com.ivoronline.springboot_endpoint_return_excel_header.controllers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

@Controller
public class MyController {

  //=======================================================================
  // GET EXCEL
  //=======================================================================
  @ResponseBody
  @GetMapping("/GetExcel")
  public String getExcel() {
    return "<a href='DownloadExcel'> Download Excel </a>";
  }

  //=======================================================================
  // DOWNLOAD EXCEL
  //=======================================================================
  @ResponseBody
  @GetMapping("/DownloadExcel")
  public ResponseEntity<StreamingResponseBody> DownloadExcel() {

    //CREATE EXCEL FILE
    Workbook workBook = new XSSFWorkbook();

    //CREATE TABLE
    Sheet    sheet = workBook.createSheet("My Sheet");
             sheet.setColumnWidth(0, 10 * 256);        //10 characters wide

    //CREATE HEADER
    Row      header = sheet.createRow(0);
             header.createCell(0).setCellValue("NAME");
             header.createCell(1).setCellValue("AGE");

    //CREATE ROW 1
    Row      row1 = sheet.createRow(1);
             row1.createCell(0).setCellValue("John");
             row1.createCell(1).setCellValue("20");

    //CREATE ROW 2
    Row      row2 = sheet.createRow(2);
             row2.createCell(0).setCellValue("Bill");
             row2.createCell(1).setCellValue("50");

    //DOWNLOAD EXCEL
    return ResponseEntity
      .ok()
      .contentType(MediaType.APPLICATION_OCTET_STREAM)
      .header(HttpHeaders.CONTENT_DISPOSITION, "inline;filename=\"People.xlsx\"")
      .body(workBook::write);
  }
}

