
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPoiReaderUtils {
    public Workbook getWorkbook(String filePath) {
        Workbook workbook = null;
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                workbook = new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public List<Sheet> getSheetList(Workbook workbook){
        List<Sheet> sheetList=new ArrayList<>();
        if (workbook==null) return sheetList;
        for (int i=0;i<workbook.getNumberOfSheets();i++){
            sheetList.add(workbook.getSheetAt(i));
        }
        return sheetList;
    }


    public List<Person> gePersonList(Sheet sheet) {
        List<Person> personList = new ArrayList<>();
        if (sheet == null) return personList;
        Row row;
        Cell cell;
        String cellData;
        // 获取最大行数
        int rowMaxNum = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rowMaxNum; i++) {
            row = sheet.getRow(i);
            if (row == null) break;
            cell = row.getCell(2);
            if (cell == null) break;
            cellData = getCellFormatValue(cell);
            if (!cellData.isEmpty() && (cellData.equals("男") || cellData.equals("女"))) {//通过性别判断是否是有效数据
                Person person = new Person();
                person.setSex(cellData);
                person.setName(getCellFormatValue(row.getCell(1)));
                cellData = getCellFormatValue(row.getCell(4));
                if (cellData!=null&&!cellData.isEmpty()){
                    person.setAge(String.valueOf(Math.round(Double.parseDouble(cellData))));
                }else {
                    try {
                        Date date=row.getCell(3).getDateCellValue();
                        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("yyyy-MM-dd");
                        Date dateEnd=simpleDateFormat.parse("2021-11-27");
                        Calendar calendar=Calendar.getInstance();
                        calendar.setTime(date);
                        Calendar calendarEnd=Calendar.getInstance();
                        calendarEnd.setTime(dateEnd);
                        int age=0;
                        for (;calendar.before(calendarEnd);calendar.add(Calendar.YEAR,1)){
                            if (calendar.get(Calendar.YEAR)<calendarEnd.get(Calendar.YEAR)) {
                                age++;
                            }
                        }
                        person.setAge(age+"");
                    } catch (ParseException e) {
                        System.out.println("年龄为空且出生日期格式不规范，无法计算年龄!!!错误位于\""+sheet.getSheetName()+"\"的"+(i+1)+"行附近。");
                        System.exit(1);
                    }

                }
                personList.add(person);
            }
        }

        return personList;
    }

    public static String getCellFormatValue(Cell cell) {
        Object cellValue;
        if (cell != null) {
            // 判断cell类型
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    // 判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                        System.out.println(cellValue);
                    } else {
                        // 数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;

                case Cell.CELL_TYPE_STRING:
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return String.valueOf(cellValue);
    }

}
