package com.example.SpringBootExcelReader;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Service
public class ExcelService {

    public List<ExcelViewModel> creteReport(MultipartFile excel) {

        List<ExcelViewModel> excelViewModelList = new ArrayList<>();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(excel.getInputStream());
            XSSFSheet sheet = workbook.getSheetAt(0);


            for(int i=0; i<sheet.getPhysicalNumberOfRows();i++) {
                ExcelViewModel excelViewModel = new ExcelViewModel();
                XSSFRow row = sheet.getRow(i);
                if(Objects.nonNull(row)){
                    for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
                            excelViewModel.setEmail(String.valueOf(row.getCell(0)));
                            excelViewModel.setName(String.valueOf(row.getCell(1)));
                            excelViewModel.setAge(String.valueOf(row.getCell(2)));
                    }
                }
                excelViewModelList.add(excelViewModel);
            }
            excelViewModelList.remove(0);

        } catch (IOException e) {
            e.printStackTrace();
        }
        return excelViewModelList;
    }
}
