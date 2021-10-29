import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class Demo {
    public static String filePath="C:\\Users\\lenovo\\Desktop\\13.18.19组选民登记表经组长核实11.23改.xls";//选民名单表格路径，会在该路径下生成一个”原名---OK“的新表
    public static double topMargin=11;//上页边距，单位厘米
    public static double bottomMargin=11;//下页边距，单位厘米
    public static double leftMargin=8;//左页边距，单位厘米
    public static double rightMargin=8;//右页边距，单位厘米

    public static short nameFontHeightInPoints=24;//名字字体大小
    public static short nameAlignment=CellStyle.ALIGN_CENTER;//名字水平对齐方式
    public static short nameVerticalAlignment=CellStyle.VERTICAL_CENTER;//名字垂直对齐方式
    public static short nameRowHeight=1000;//名字行高

    public static short sexFontHeightInPoints=12;//性别字体大小
    public static short sexAlignment=CellStyle.ALIGN_RIGHT;//性别水平对齐方式
    public static short sexVerticalAlignment=CellStyle.VERTICAL_CENTER;//性别垂直对齐方式
    public static short sexRowHeight=500;//性别行高

    public static short ageFontHeightInPoints=6;//年龄字体大小
    public static short ageAlignment=CellStyle.ALIGN_RIGHT;//年龄水平对齐方式
    public static short ageVerticalAlignment=CellStyle.VERTICAL_CENTER;//年龄垂直对齐方式
    public static short ageRowHeight=500;//年龄行高

    public static short columnWidth =20;//默认列宽

    public static void main(String[] args) throws IOException {
        if (new File(filePath).exists()) {
            String targetFilePath=filePath.replace(".xls","---OK.xls");
            File file=new File(targetFilePath);
            file.createNewFile();
            ExcelPoiReaderUtils readerUtils = new ExcelPoiReaderUtils();
            Workbook workbook = readerUtils.getWorkbook(filePath);//根据路径获取工作簿
            Workbook targetWorkbook;
            if (targetFilePath.contains(".xlsx")) {
                targetWorkbook = new XSSFWorkbook();
            }else {
                targetWorkbook=new HSSFWorkbook();
            }
            Sheet targetSheet;
            List<Sheet> sheetList = readerUtils.getSheetList(workbook);//获取所有工作表
            List<Person> personList;
            for (Sheet sheet : sheetList) {//遍历工作表
                personList = readerUtils.gePersonList(sheet);
                targetSheet=targetWorkbook.createSheet(sheet.getSheetName());
                targetSheet.setMargin(Sheet.TopMargin,topMargin/2.54);
                targetSheet.setMargin(Sheet.BottomMargin,bottomMargin/2.54);
                targetSheet.setMargin(Sheet.LeftMargin,leftMargin/2.54);
                targetSheet.setMargin(Sheet.RightMargin,rightMargin/2.54);
                new ExcelPoiWriterUtils().writeExcel(targetWorkbook,targetSheet, personList);//写数据
            }
            try {
                OutputStream out = new FileOutputStream(targetFilePath);
                targetWorkbook.write(out);
                out.close();
            } catch (Exception e){
                e.printStackTrace();
            }
        }else {
            System.out.println("文件不存在，请检查路径");
        }
    }


}
