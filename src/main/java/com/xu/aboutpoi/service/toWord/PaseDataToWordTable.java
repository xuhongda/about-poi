package com.xu.aboutpoi.service.toWord;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.List;

/**
 * <p>
 * data 数据 转换生成 word table
 * </p>
 *
 * @author xuhongda on 2018/9/17
 * com.xu.aboutpoi.service.toExcel
 * about-poi
 */
public class PaseDataToWordTable {
    protected static int rows;
    protected static int columns;

    public static <T> void newWordDoc(Class<T> tClass, List<T> objList, String filename, Integer line) throws Exception {
        //构建文档对象
        XWPFDocument document = createDoc(tClass);

        //新建一个段落
        XWPFParagraph p = document.createParagraph();

        paragraph(p, document);

        //构建表格
        XWPFTable xTable = createTable(tClass, objList, document, line);
        //设置表头数据
        createTableHeader(tClass, xTable);
        //插入正文数据
        insertContext(tClass, xTable, objList, line);

        //写入文件
        String separator = File.separator;
        FileOutputStream fos = new FileOutputStream(new File("C:" + separator + "java&wordTest", filename + ".docx"));
        document.write(fos);
        fos.close();
    }

    protected static void paragraph(XWPFParagraph p, XWPFDocument document) {
        // 设置段落的对齐方式
        p.setAlignment(ParagraphAlignment.CENTER);
        //设置下边框
        p.setBorderBottom(Borders.SINGLE);
        //设置上边框
        p.setBorderTop(Borders.SINGLE);
        //设置右边框
        p.setBorderRight(Borders.DOUBLE);
        //设置左边框
        p.setBorderLeft(Borders.DOUBLE);
        //创建段落文本
        XWPFRun r = p.createRun();
        r.setText("这是POI创建的一个Word段落文本");
        //设置为粗体
        r.setBold(true);
        //设置颜色
        r.setColor("EEC591");
        // 新建一个段落
        p = document.createParagraph();
        r = p.createRun();
        r.setText("亿达中国控股有限公司（股票代码：3639.HK）是中国领先的商务园区运营商，于2014年6月27日在香港联交所主板上市，主要业务涵盖商务园区运营，房地产综合开发，建筑，装修，园林，物业管理等产业，就已竣工的建筑面积而言，亿达中国是中国最大的商务园区开发商。作为商务园区运营专家，自1998年开始，亿达中国先后开发和运营了大连软件园、大连生态科技创新城...");
    }


    /**
     * 创建文档对象
     *
     * @param tClass
     * @param <T>
     * @return
     */
    protected static <T> XWPFDocument createDoc(Class<T> tClass) {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph title = document.createParagraph();
        //设置标题对齐方式
        title.setAlignment(ParagraphAlignment.CENTER);
        title.setVerticalAlignment(TextAlignment.TOP);
        XWPFRun xr = title.createRun();
        //设置字体是否加粗
        xr.setBold(true);
        xr.setFontSize(20);
        String name = tClass.getName();
        xr.setText(name.substring(name.lastIndexOf(".") + 1) + "表");
        return document;
    }


    /**
     * 构建表格结构
     *
     * @param tClass
     * @param objList
     * @param xd
     * @param <T>
     * @param line    表头行数
     * @return
     */
    protected static <T> XWPFTable createTable(Class<T> tClass, List<T> objList, XWPFDocument xd, Integer line) {
        //设置表格的行数和列数
        tableRowsColumns(tClass, objList);
        //需要多构造line行，用于表头
        XWPFTable xTable = xd.createTable(rows + line, columns);

        //表格属性
        CTTblPr tblPr = xTable.getCTTbl().getTblPr();
        tblPr.getTblW().setType(STTblWidth.DXA);
        tblPr.getTblW().setW(BigInteger.valueOf(8300));

        // 设置上下左右四个方向的距离，可以将表格撑大
        xTable.setCellMargins(100, 100, 100, 100);

        // 表格
        List<XWPFTableCell> tableCells = xTable.getRow(0).getTableCells();
        XWPFTableCell cell = tableCells.get(0);

        // 相当于合并单元格
        // XWPFParagraph newPara = new XWPFParagraph(cell.getCTTc().addNewP(), cell);
        // XWPFRun run = newPara.createRun();
        // 内容居中显示
        // newPara.setAlignment(ParagraphAlignment.CENTER);

        return xTable;
    }

    /**
     * 创建表头
     *
     * @param tClass
     * @param xTable
     * @param <T>
     */
    protected static <T> void createTableHeader(Class<T> tClass, XWPFTable xTable) {
        //获取字段集合
        Field[] fields = tClass.getDeclaredFields();
        //获取表格第一行
        XWPFTableRow xRow = xTable.getRow(0);
        XWPFTableRow row2 = xTable.getRow(1);


        //填充表头信息。
        for (int i = 0; i < fields.length; ++i) {
            xRow.getCell(i).setText(fields[i].getName());
        }
        row2.getCell(0).setText("易达");
        row2.getCell(1).setText("云图");
        row2.getCell(2).setText("园区");
        row2.getCell(3).setText("xxx");
    }

    /**
     * 插入数据库正文
     *
     * @param tClass
     * @param xTable
     * @param objList
     * @param <T>
     */
    protected static <T> void insertContext(Class<T> tClass, XWPFTable xTable, List<T> objList, Integer line) {
        tableRowsColumns(tClass, objList);
        Field[] fields = tClass.getDeclaredFields();

        try {
            XWPFTableRow xRow;
            //处理每一行
            for (int i = 0; i < rows - 1; ++i) {
                //获取行从1 到 rows - 2(包含rows-2);
                xRow = xTable.getRow(i + line);
                //从0到rows - 2,最后一行单独处理
                T obj = objList.get(i);
                for (int j = 0; j < columns; ++j) {
                    //写每一列数据
                    fields[j].setAccessible(true);
                    xRow.getCell(j).setText(fields[j].get(obj) + "");
                }
            }

            // 单独处理最后一行
            T obj = objList.get(rows - 1);
            //获取最后一行
            xRow = xTable.getRow(rows + line - 1);
            for (int j = 0; j < columns; ++j) {
                xRow.getCell(j).setText(fields[j].get(obj) + "");
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }


    /**
     * 获取行数和列数
     *
     * @param tClass
     * @param objList
     * @param <T>
     */
    protected static <T> void tableRowsColumns(Class<T> tClass, List<T> objList) {
        Field[] fields = tClass.getDeclaredFields();
        rows = objList.size();
        columns = fields.length+1;
    }

    //设置表格风格
//    protected
}
