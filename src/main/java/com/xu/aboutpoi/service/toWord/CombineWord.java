package com.xu.aboutpoi.service.toWord;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author xuhongda on 2018/9/17
 * com.xu.aboutpoi.service.toWord
 * about-poi
 */
public class CombineWord {



    public static<T> void getTabele(Class<T> tClass, List<T> objList, String filename, Integer line) throws IOException {
        XWPFDocument document = PaseDataToWordTable.createDoc(tClass);

        XWPFTable xwpfTable = PaseDataToWordTable.createTable(tClass,objList,document,line);
        //跨列操作
        mergeCellsHorizontal(xwpfTable,0,2,4);
        //跨行操作
        mergeCellsVertically(xwpfTable,0,1,2);

        //设置表头数据
        PaseDataToWordTable.createTableHeader(tClass, xwpfTable);
        //插入正文数据
        PaseDataToWordTable.insertContext(tClass, xwpfTable, objList, line);

        //写入文件
        String separator = File.separator;
        FileOutputStream fos = new FileOutputStream(new File("C:" + separator + "java&wordTest", filename + ".docx"));
        document.write(fos);
        fos.close();
    }

    /**
     *
     * @param table
     * @param row 第几行从0开始
     * @param fromCell 横跨几行
     * @param toCell 从第几列开始
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if ( cellIndex == fromCell ) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * word跨行并单元格
     * @param table
     * @param col
     * @param fromRow
     * @param toRow
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if ( rowIndex == fromRow ) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**/


}
