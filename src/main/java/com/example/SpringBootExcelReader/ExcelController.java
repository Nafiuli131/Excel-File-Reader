package com.example.SpringBootExcelReader;

import java.io.IOException;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ExcelController {

	@Autowired
	ExcelService excelService;

	@PostMapping("excel")
	public List<ExcelViewModel> excelReader(@RequestParam("file") MultipartFile excel) {

		return excelService.creteReport(excel);

	}
}
