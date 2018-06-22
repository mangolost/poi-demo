package mangolost.poistudy.service;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 *
 */
public class ExcelService {

    /**
     * 创建工作簿和工作表
     *
     * @param pathName
     */
    public void createWorkbook(String pathName) throws IOException {

        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(new File(pathName)); //要创建的文件流

        //Create Blank workbook
        Workbook workbook = new XSSFWorkbook(); //新建一个空白工作簿
        Sheet sheet1 = workbook.createSheet(); //创建工作表1，注意至少要创建一个sheet，否则创建的工作簿打开会报错
        workbook.setSheetName(0, "Sheet reainll"); //如果不设置名称，按照创建顺序默认为Sheet0、Sheet1、...
        Sheet sheet2 = workbook.createSheet(); //创建工作表2
        workbook.setSheetName(1, "Sheet mangolost");

        workbook.write(out); //把工作簿写入流，生成xlsx文件
        out.close(); //关闭文档流
        System.out.println("create workbook and written successfully");
    }

    /**
     * 读取工作表,并按照行和列顺序组成一个二维list
     *
     * @param pathName
     */
    public List<List<Object>> readWorkbook(String pathName) throws IOException {

        File file = new File(pathName);
        Workbook workbook = new XSSFWorkbook(new FileInputStream(file)); //读取工作簿
        // 读取第一张表格内容
        Sheet sheet = workbook.getSheetAt(0);
        List<List<Object>> list = new LinkedList<>(); //待获取数据的list
        int firstRowIndex = sheet.getFirstRowNum(); //指第一个出现非空cell的行的index
        int lastRowIndex = sheet.getLastRowNum(); //指最后一个出现非空cell的行的index
        for (int i = firstRowIndex; i <= lastRowIndex; i++) {
            Row xssfRow = sheet.getRow(i); //当前行
            if (xssfRow == null) {
                continue; //这里的策略是：如果此行为空，则跳过此行
            }
            List<Object> rowList = new ArrayList<>();
            int firstCellIndex = xssfRow.getFirstCellNum(); //指当前行中第一个出现非空cell的列的index
            int lastCellIndex = xssfRow.getLastCellNum(); //指当前行中最后一个出现非空cell的列的index
            for (int j = firstCellIndex; j <= lastCellIndex; j++) {
                Cell cell = xssfRow.getCell(j); //当前单元格
                if (cell == null) {
                    continue;
                }
                Object value = null;
                switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_STRING:
                        //String类型返回String数据
                        value = cell.getStringCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        //日期数据返回Timestamp类型
                        if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
                            value = new Timestamp(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).getTime());
                        } else {
                            //数值类型返回double类型的数字
                            //System.out.println(cell.getNumericCellValue()+":格式："+cell.getCellStyle().getDataFormatString());
                            value = cell.getNumericCellValue(); //这里，整数也会变成double， 比如33会变成33.0
                        }
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        //布尔类型
                        value = cell.getBooleanCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_BLANK:
//                        value = null; //如果没有内容，默认为null
                        break;
                    default:
                        value = cell.toString();
                }
                rowList.add(value);
            }
            if (rowList.size() != 0) {
                list.add(rowList);
            }
        }
        return list;
    }

    /**
     * 编辑、更新工作簿
     * @param pathName
     */
    public void updateWorkbook(String pathName) throws IOException {
        File file = new File(pathName);
        Workbook wookbook = new XSSFWorkbook(new FileInputStream(file)); //读取工作簿
        // 读取第一张表格内容
        Sheet sheet = wookbook.getSheetAt(0);

        //把第7行，第3列的单元格数值设置为777
        Row row = sheet.getRow(6);
        if (row == null) {
            row = sheet.createRow(6);
        }
        Cell cell = row.createCell(2);
        cell.setCellValue(777);

        //设置单元格样式
        CellStyle cellStyle = wookbook.createCellStyle();

        //设置字体
        Font font = wookbook.createFont();
        font.setBold(true);
        font.setItalic(true);
        font.setFontName("微软雅黑");
        font.setColor(HSSFColor.RED.index);
        cellStyle.setFont(font);

        //设置背景颜色
        cellStyle.setFillBackgroundColor(HSSFColor.BLUE.index);

        cell.setCellStyle(cellStyle);

        FileOutputStream out = new FileOutputStream(file);
        wookbook.write(out); //把工作簿写入流，更新xlsx文件
        out.close(); //关闭文档流
        System.out.println("update and save workbook successfully");
    }

    /**
     * 删除工作簿，这个其实是java文件操作，跟poi没关系了
     * @param pathName
     */
    public void deleteWorkbook(String pathName) {
        File file = new File(pathName);
        // 如果文件路径所对应的文件存在，并且是一个文件，则直接删除
        if (file.exists() && file.isFile()) {
            if (file.delete()) {
                System.out.println("删除文件" + pathName + "成功！");
            } else {
                System.out.println("删除文件失败！");
            }
        } else {
            System.out.println("删除文件失败：" + pathName + "不存在！");
        }
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        ExcelService excelService = new ExcelService();

        excelService.deleteWorkbook("mango.xlsx");
        excelService.createWorkbook("mango.xlsx");

        List<List<Object>> list = excelService.readWorkbook("mango.xlsx");
        System.out.println("aaaa");

        excelService.updateWorkbook("mango.xlsx");


        excelService.deleteWorkbook("ccc.xlsx");
        excelService.createWorkbook("ddd.xlsx");
        Thread.sleep(5000);
        excelService.deleteWorkbook("ddd.xlsx");

    }
}
