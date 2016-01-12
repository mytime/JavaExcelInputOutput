package com.niupei.main;

import com.niupei.bean.Book;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Excle 导入导出
 */
public class ExcleBook {

    //导出方法
    public void excleOut() {
        //定义WritableWorkbook类型的对象，带表Excle对象
        WritableWorkbook book = null;

        try {
            //创建excle对象
            book = Workbook.createWorkbook(new File("book.xls"));
            //通过excle对象创建一个选项卡对象
            WritableSheet sheet = book.createSheet("sheet1", 0);
            //创建单元格对象，参数：列 行 值
            Label la = new Label(0, 2, "test");
            //将创建好的单元格对象放入选项卡中
            sheet.addCell(la);
            //输出文件到目标路径
            book.write();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } finally {
            //由于book 也是数据流，需要关闭操作
            try {
                book.close();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }


    public static void main(String[] args) {
        //实例化
        ExcleBook eb = new ExcleBook();
        //导出
        eb.excleOut();
    }


}
