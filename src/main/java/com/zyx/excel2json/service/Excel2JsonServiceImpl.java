package com.zyx.excel2json.service;

import com.zyx.excel2json.common.ExcelUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;


@Service
public class Excel2JsonServiceImpl implements Excel2JsonService {

    @Override
    public void excel2JSon(MultipartFile file) {

        try {
            //Excel组件
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
            // 获取sheet页数
            int sheetNum = workbook.getNumberOfSheets();
            LinkedHashMap<String,List<LinkedHashMap<String,String>>> json = new LinkedHashMap<>(sheetNum);
            for (int i = 0; i < sheetNum; i++) {
                String sheetName = workbook.getSheetName(i);
                XSSFSheet sheet = workbook.getSheetAt(i);
                int rowNum = sheet.getLastRowNum();
                List<LinkedHashMap<String,String>> rowList = new ArrayList<>(rowNum);
                XSSFRow titleRow = sheet.getRow(0);
                XSSFRow row;
                for (int j = 1; j < rowNum; j++) {
                    row = sheet.getRow(j);
                    int columnNum = row.getLastCellNum();
                    LinkedHashMap<String,String > rowMap = new LinkedHashMap<>(columnNum);
                    for (int k = 0; k < columnNum; k++) {
                        String title = ExcelUtil.doGetCellValue(titleRow.getCell(k));
                        String value = ExcelUtil.doGetCellValue(row.getCell(k));
                        rowMap.put(title,value);
                    }
                    rowList.add(rowMap);
                }
                json.put(sheetName,rowList);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
