package com.zyx.excel2json.common;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

public class ExcelUtil {

    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static String doGetCellValue(XSSFCell cell){
        String value = null;
        if(null != cell){
            CellType cellType = cell.getCellType();
            switch (cellType){
                case NUMERIC:
                    //数字格式包括日期，需要特别处理
                    if(DateUtil.isCellDateFormatted(cell)){
                        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                        LocalDate date = cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                        value = dateTimeFormatter.format(date);
                    }else{
                        value = String.valueOf((Double) cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN: value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA: value = cell.getCellFormula();
                    break;
                case ERROR: value = "ERROR";
                    break;
                case BLANK: value = null;
                    break;
                default:value = cell.getStringCellValue();
                    break;
            }
        }
        return value;
    }
}
