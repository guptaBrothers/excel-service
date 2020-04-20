package com.amdocs.document.excel.reader;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.amdocs.document.excel.model.UserSheet;

@Component
public class ExcelFileReader extends AbstractExcelReader<UserSheet>
{
    @Value("${file.path}")
    private String path;

    @Override public List<UserSheet> covert(String fileName) throws IOException, InvalidFormatException
    {
        this.setSourceFile(fileName);
        List<UserSheet> list = new ArrayList<>();

        Workbook wb;

        File file = new File(path  + this.getSourceFile());

            wb = WorkbookFactory.create(file);


        this.setNumberOfSheets(wb.getNumberOfSheets());


        for (int i = 0; i < this.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            UserSheet userSheet = new UserSheet();
            userSheet.setSheetName(sheet.getSheetName());
            List<Map<Integer , Object>> rows = new ArrayList<>();

            for(int j=sheet.getFirstRowNum(); j<=sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if(row==null) {
                    continue;
                }
                boolean hasValues = false;
                Map<Integer , Object> colums = new HashMap<>();
                for(int k=0; k<=row.getLastCellNum(); k++) {
                    Cell cell = row.getCell(k);
                    if(cell!=null) {
                        Object value = cellToObject(cell);
                        colums.put(cell.getColumnIndex() , value);
                    } else {
                       // colums.put(cell.getColumnIndex() , null);
                    }
                }

                rows.add(colums);
            }

            userSheet.setRows(rows);
            list.add(userSheet);
        }

        return list;
    }

    private Object cellToObject(Cell cell) {

        int type = cell.getCellType();

        if(type == Cell.CELL_TYPE_STRING) {
            return cleanString(cell.getStringCellValue());
        }

        if(type == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getBooleanCellValue();
        }

        if(type == Cell.CELL_TYPE_NUMERIC) {

            if (cell.getCellStyle().getDataFormatString().contains("%")) {
                return cell.getNumericCellValue() * 100;
            }

            return numeric(cell);
        }

        if(type == Cell.CELL_TYPE_FORMULA) {
            switch(cell.getCachedFormulaResultType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return numeric(cell);
                case Cell.CELL_TYPE_STRING:
                    return cleanString(cell.getRichStringCellValue().toString());
            }
        }

        return null;

    }

    private String cleanString(String str) {
        return str.replace("\n", "").replace("\r", "");
    }

    private Object numeric(Cell cell) {
        if(HSSFDateUtil.isCellDateFormatted(cell)) {

            return cell.getDateCellValue();
        }
        return cell.getNumericCellValue();
    }
}
