import org.apache.poi.ss.usermodel.*;

import java.util.*;

public class ExcelPoiWriterUtils {
    public void writeExcel(Workbook workbook, Sheet sheet, List<Person> personList) {
        sheet.setDefaultColumnWidth(Demo.columnWidth);
        Row row;

        Font fontName = workbook.createFont();
        fontName.setFontHeightInPoints(Demo.nameFontHeightInPoints);
        CellStyle cellStyleName = workbook.createCellStyle();
        cellStyleName.setFont(fontName);
        cellStyleName.setAlignment(Demo.nameAlignment);
        cellStyleName.setVerticalAlignment(Demo.nameVerticalAlignment);

        Font fontSex = workbook.createFont();
        fontSex.setFontHeightInPoints(Demo.sexFontHeightInPoints);
        CellStyle cellStyleSex = workbook.createCellStyle();
        cellStyleSex.setFont(fontSex);
        cellStyleSex.setAlignment(Demo.sexAlignment);
        cellStyleName.setVerticalAlignment(Demo.sexVerticalAlignment);

        Font fontAge = workbook.createFont();
        fontAge.setFontHeightInPoints(Demo.ageFontHeightInPoints);
        CellStyle cellStyleAge = workbook.createCellStyle();
        cellStyleAge.setFont(fontAge);
        cellStyleAge.setAlignment(Demo.ageAlignment);
        cellStyleName.setVerticalAlignment(Demo.ageVerticalAlignment);

        Cell cellName,cellAge,cellSex;
        for (int i = 0; i < personList.size(); i++) {
            Person person = personList.get(i);

            row=sheet.createRow(i*3);
            row.setHeight(Demo.nameRowHeight);
            cellName=row.createCell(0);
            cellName.setCellValue(person.getName());
            cellName.setCellStyle(cellStyleName);

            row=sheet.createRow(i*3+1);
            row.setHeight(Demo.sexRowHeight);
            cellSex= row.createCell(0);
            cellSex.setCellValue(person.getSex());
            cellSex.setCellStyle(cellStyleSex);

            row=sheet.createRow(i*3+2);
            row.setHeight(Demo.ageRowHeight);
            cellAge=row.createCell(0);
            cellAge.setCellValue(person.getAge());
            cellAge.setCellStyle(cellStyleAge);
//            row = sheet.createRow(i * 5);
//            row.createCell(0);
//            row.createCell(1);
//            //合并单元格
//            CellRangeAddress rangeAddress = new CellRangeAddress(i * 5, i * 5, 0, 1);
//            sheet.addMergedRegion(rangeAddress);
//            Cell cell0 = row.getCell(0);
//            cell0.setCellValue(person.getName());
//            cell0.setCellStyle(cellStyleName);
//
//            row = sheet.createRow(i * 5 + 4);
//
//            Cell cell1 = row.createCell(0);
//            cell1.setCellValue(person.getSex());
//            cell1.setCellStyle(cellStyleSex);
//
//            Cell cell2 = row.createCell(1);
//            cell2.setCellValue(person.getAge());
//            cell2.setCellStyle(cellStyleAge);
        }
    }

}
