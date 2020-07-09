package tangcloud;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExeclWriteTest {

    String PATH = "E:/ideaWorkspace/easyexcel/";
    @Test
    public void testWrite03() throws IOException {

        //生成execl-03版
        //1. 创建一个工作簿
        // 03版本
        //Workbook workbook = new HSSFWorkbook();
        // 07 版本
        Workbook workbook = new XSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("Execl(03版本)代码学习");
        //3. 创建一行 (1 ,1)
        Row row = sheet.createRow(0);
        //4. 创建一个单元格
        Cell cell = row.createCell(0);
        cell.setCellValue("第一行第一个单元格");
        // 1,2
        Cell cell2 = row.createCell(1);
        cell2.setCellValue("第一行第二个单元格");
        //4 创建第二行
        Row row2 = sheet.createRow(1);
        // 创建第二行第一列单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("第二行第一个单元格");
        // 创建第二行第二列单元格
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);
        //生成一张表 03 版本excel 结尾是 xls
        //FileOutputStream fileOutputStream =  new FileOutputStream(PATH + "Execl03版本代码测试"+ ".xls");
        FileOutputStream fileOutputStream =  new FileOutputStream(PATH + "Execl07版本代码测试"+ ".xlsx");
        //输出到文件中
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("生成完毕!!");

    }
    @Test
    public void testWrite07() throws IOException {
        //生成execl-03版
        //1. 创建一个工作簿
        // 07 版本
        Workbook workbook = new XSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("Execl(03版本)代码学习");
        //3. 创建一行 (1 ,1)
        Row row = sheet.createRow(0);
        //4. 创建一个单元格
        Cell cell = row.createCell(0);
        cell.setCellValue("第一行第一个单元格");
        // 1,2
        Cell cell2 = row.createCell(1);
        cell2.setCellValue("第一行第二个单元格");
        //4 创建第二行
        Row row2 = sheet.createRow(1);
        // 创建第二行第一列单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("第二行第一个单元格");
        // 创建第二行第二列单元格
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);
        //生成一张表 07 版本excel 结尾是 xlxs
        FileOutputStream fileOutputStream =  new FileOutputStream(PATH + "Execl07版本代码测试"+ ".xlsx");
        //输出到文件中
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("生成完毕!!");
    }

    /**
     * 测试写入时间 (03版本最多存储65536行数据 写入速度慢)
     */
    @Test
    public void testWrite03BigTime() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();
        //创建一个簿
        Workbook workbook = new HSSFWorkbook();
        //创建一个表
        Sheet sheet = workbook.createSheet("测试03版本写入65536行数据的时间");
        //写入数据
        for (int rowNum = 0; rowNum < 65536;rowNum++){
            //循环创建行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("循环写入完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试03版本写入时间"+ ".xls");
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        //记录结束时间
        long end = System.currentTimeMillis();
        //获取生成时间
        System.out.println((double)(end-begin)/1000);
    }

    /**
     * 测试写入时间 (07版本没有数量限制 但是写入速度慢)
     */
    @Test
    public void testWrite07BigTime() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();
        //创建一个簿
        Workbook workbook = new XSSFWorkbook();
        //创建一个表
        Sheet sheet = workbook.createSheet("测试07版本写入65536行数据的时间");
        //写入数据
        for (int rowNum = 0; rowNum < 65536;rowNum++){
            //循环创建行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("循环写入完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试07版本写入时间"+ ".xlsx");
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        //记录结束时间
        long end = System.currentTimeMillis();
        //获取生成时间
        System.out.println((double)(end-begin)/1000);
    }

    /**
     * 提高测试写入时间 ( 07版本没有数量限制但是写入速度慢)
     * 会生成临时文件
     */
    @Test
    public void testWrite07BigTimeS() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();
        //创建一个簿
        Workbook workbook = new SXSSFWorkbook();
        //创建一个表
        Sheet sheet = workbook.createSheet("测试07版本写入65536行数据的时间");
        //写入数据
        for (int rowNum = 0; rowNum < 65536;rowNum++){
            //循环创建行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("循环写入完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试07版本写入时间"+ ".xlsx");
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook) workbook).dispose();
        //记录结束时间
        long end = System.currentTimeMillis();
        //获取生成时间
        System.out.println((double)(end-begin)/1000);
    }

}
