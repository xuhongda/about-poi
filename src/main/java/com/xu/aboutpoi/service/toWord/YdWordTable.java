package com.xu.aboutpoi.service.toWord;

import com.xu.aboutpoi.entity.Customer;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;


/**
 * @author xuhongda on 2018/9/18
 * com.xu.aboutpoi.service.toWord
 * about-poi
 */
public class YdWordTable {
    /**
     * 行
     */
    private static Integer rows;
    /**
     * 列
     */
    private static Integer columns;


    private static <T> void tableRowsColumns(List<T> objList) {
        rows = objList.size() + 4;
        columns = 8;
    }

    /**
     * 创建文档对象
     * 设置表格样式
     *
     * @return
     */
    protected static XWPFDocument createDoc() {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph title = document.createParagraph();
        //设置标题对齐方式
        title.setAlignment(ParagraphAlignment.CENTER);
        title.setVerticalAlignment(TextAlignment.TOP);
        XWPFRun xr = title.createRun();
        //字体设置
        xr.setBold(true);
        xr.setFontSize(20);
        xr.setColor("F08080");
        xr.setText("租金物业费支付明细表");
        return document;
    }

    /**
     * 创建表格
     * 行列数目
     *
     * @param xd
     * @param list
     * @return
     */
    protected static XWPFTable createTable(XWPFDocument xd, List list) {
        tableRowsColumns(list);
        //构造表格（行x列） 需要多构造line行，用于表头
        XWPFTable xTable = xd.createTable(rows, columns);
        //表格属性
        CTTblPr tblPr = xTable.getCTTbl().getTblPr();
        tblPr.getTblW().setType(STTblWidth.DXA);
        tblPr.getTblW().setW(BigInteger.valueOf(8300));
        // 设置上下左右四个方向的距离，可以将表格撑大
        xTable.setCellMargins(100, 100, 100, 100);
        return xTable;
    }

    /**
     * 创建表头
     *
     * @param xTable
     */
    protected static <T> void createTableHeader(XWPFTable xTable, List<T> list) {

        //获取表格第一行
        XWPFTableRow row1 = xTable.getRow(0);
        XWPFTableRow row2 = xTable.getRow(1);

        //设置表头颜色
        for (int i = 0; i < 8; i++) {
            XWPFTableCell cell1 = row1.getCell(i);
            XWPFTableCell cell2 = row2.getCell(i);
            cell1.setColor("63B8FF");
            cell2.setColor("63B8FF");
        }

        //设置表头行
        row1.getCell(0).setText("项目");
        row1.getCell(1).setText("款期");
        row1.getCell(2).setText("应付截止日期");
        row1.getCell(3).setText("代表期限");
        //row1.getCell(4).setText("期限");
        row1.getCell(5).setText("应付租金");
        row1.getCell(6).setText("应付物业费");
        row1.getCell(7).setText("应付费用小计");

        row2.getCell(3).setText("始");
        row2.getCell(4).setText("止");

        //设置列
        xTable.getRow(2).getCell(0).setColor("CDC9C9");
        xTable.getRow(2).getCell(0).setText("履约保证金");

        xTable.getRow(3).getCell(0).setColor("CDBA96");
        xTable.getRow(3).getCell(0).setText("租金物业费");

        xTable.getRow(rows - 1).getCell(0).setColor("CD919E");
        xTable.getRow(rows - 1).getCell(0).setText("合计");
    }


    public static void main(String[] args) throws IOException {
        List<Customer> list = getList();
        XWPFDocument document = createDoc();
        XWPFTable table = createTable(document, list);
        //跨列
        CombineWord.mergeCellsHorizontal(table, 0, 3, 4);
        //跨行
        CombineWord.mergeCellsVertically(table, 0, 0, 1);
        CombineWord.mergeCellsVertically(table, 1, 0, 1);
        CombineWord.mergeCellsVertically(table, 2, 0, 1);
        CombineWord.mergeCellsVertically(table, 5, 0, 1);
        CombineWord.mergeCellsVertically(table, 6, 0, 1);
        CombineWord.mergeCellsVertically(table, 7, 0, 1);
        //租金物业费
        CombineWord.mergeCellsVertically(table, 0, 3, list.size() + 2);
        //设置表头
        createTableHeader(table, list);
        //写入文件
        String separator = File.separator;
        FileOutputStream fos = new FileOutputStream(new File("C:" + separator + "java&wordTest", "2018-新租-深圳产业软件园测试-环球金融集团-1010订单信息" + ".docx"));
        document.write(fos);
        fos.close();
    }

    static List<Customer> getList() {

        for (int i = 0; i < 100; i++) {
            Customer customer = new Customer();
            customer.setContact("xxx" + i + i % 2);
            customer.setEmail("xx@c.com");
            customer.setId(Integer.valueOf(i).longValue());
            customer.setName("cc" + i);
            customer.setRemark("100" + i);
            customer.setTelephone("131" + i + i % 2);
        }
        Customer customer1 = new Customer();
        Customer customer2 = new Customer();
        Customer customer3 = new Customer();
        Customer customer4 = new Customer();
        Customer customer5 = new Customer();
        customer1.setContact("xxx");
        customer1.setEmail("xx@c.com");
        customer1.setId(1L);
        customer1.setName("cc1");
        customer1.setRemark("100|");
        customer1.setTelephone("131");
        customer2.setContact("xxx");
        customer2.setEmail("xx@c.com");
        customer2.setId(2L);
        customer2.setName("cc2");
        customer2.setRemark("200");
        customer2.setTelephone("131");
        customer3.setContact("xxx");
        customer3.setEmail("xx@c.com");
        customer3.setId(3L);
        customer3.setName("cc3");
        customer3.setRemark("300");
        customer3.setTelephone("131");
        customer4.setContact("xxx");
        customer4.setEmail("xx@c.com");
        customer4.setId(3L);
        customer4.setName("cc3");
        customer4.setRemark("300");
        customer4.setTelephone("131");
        customer5.setContact("xxx");
        customer5.setEmail("xx@c.com");
        customer5.setId(3L);
        customer5.setName("cc3");
        customer5.setRemark("300");
        customer5.setTelephone("131");
        List<Customer> customers = Arrays.asList(customer1, customer2, customer3, customer4, customer5);
        return customers;
    }
}
