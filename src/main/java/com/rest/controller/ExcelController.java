package com.rest.controller;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ExcelController {
	
	
	 
	 @PostMapping("/upload")
	    public ResponseEntity<List<List<String>>> uploadExcelFile(@RequestParam("file") MultipartFile file) {
	        try {
	            Workbook workbook = new XSSFWorkbook(file.getInputStream());
	            Sheet sheet = workbook.getSheetAt(0);
	            Iterator<Row> rows = sheet.iterator();

	            List<List<String>> data = new ArrayList<>();

	            while (rows.hasNext()) {
	                Row row = rows.next();
	                Iterator<Cell> cells = row.cellIterator();
	                List<String> rowData = new ArrayList<>();

	                while (cells.hasNext()) {
	                    Cell cell = cells.next();
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            rowData.add(cell.getStringCellValue());
	                            break;
	                        case NUMERIC:
	                            rowData.add(Double.toString(cell.getNumericCellValue()));
	                            break;
	                        case BOOLEAN:
	                            rowData.add(Boolean.toString(cell.getBooleanCellValue()));
	                            break;
	                        default:
	                            rowData.add("");
	                    }
	                }
	                data.add(rowData);
	            }
	            workbook.close();
	            return new ResponseEntity<>(data, HttpStatus.OK);
	        } catch (IOException e) {
	            e.printStackTrace();
	            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
	        }
	    }
}
