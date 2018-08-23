/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package schedule;

import java.io.File;
import java.util.ArrayList;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.format.Colour;
import jxl.write.WriteException;

/**
 *
 * @author wgyklh
 */
public class Schedule {

    /**
     * @param args the command line arguments
     * @throws java.lang.Exception
     */
    public static void main(String[] args) throws Exception {

        //实例化
        Schedule scd = new Schedule();

        //创建文件读取总课表
        File sch1 = new File("C:\\Users\\Lenovo\\Desktop\\1711003.xls");
        Workbook wb1 = Workbook.getWorkbook(sch1);
        Sheet sheet1 = wb1.getSheet(0);

        //创建文件存储分周课表
        File sch2 = new File("C:\\Users\\Lenovo\\Desktop\\1711003(devided).xls");
        WritableWorkbook wb2 = Workbook.createWorkbook(sch2);

        //设置标题格式
        WritableFont font1 = new WritableFont(WritableFont.createFont("微软雅黑"), 20, WritableFont.BOLD);
        font1.setColour(Colour.WHITE);
        WritableCellFormat format1 = new WritableCellFormat(font1);
        format1.setWrap(true);
        format1.setAlignment(jxl.format.Alignment.CENTRE);
        format1.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THICK, Colour.WHITE);

        //设置正文格式
        WritableFont font2 = new WritableFont(WritableFont.createFont("仿宋"), 12, WritableFont.BOLD);
        WritableCellFormat format2 = new WritableCellFormat(font2);
        format2.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.MEDIUM_DASH_DOT_DOT, Colour.WHITE);

        //自定义修改颜色并设置背景色
        wb2.setColourRGB(Colour.LIGHT_BLUE, 215, 255, 255);
        wb2.setColourRGB(Colour.DARK_BLUE, 48, 84, 150);
        wb2.setColourRGB(Colour.DARK_BLUE2, 189, 215, 238);
        wb2.setColourRGB(Colour.DARK_YELLOW, 255, 217, 102);
        wb2.setColourRGB(Colour.YELLOW2, 255, 242, 204);
        format1.setBackground(Colour.DARK_BLUE);
        format2.setBackground(Colour.LIGHT_BLUE);

        for (int week = 1; week < 20; week++) {
            //创建sheet
            WritableSheet sheet = wb2.createSheet("第" + week + "周", week - 1);
            //添加表头、标题
            sheet.addCell(new Label(0, 0, "第" + week + "周课表", format1));
            sheet.addCell(new Label(0, 1, "", scd.SetHead(0)));
            sheet.addCell(new Label(1, 1, "", scd.SetHead(0)));
            sheet.addCell(new Label(2, 1, "星期一", scd.SetHead(0)));
            sheet.addCell(new Label(3, 1, "星期二", scd.SetHead(0)));
            sheet.addCell(new Label(4, 1, "星期三", scd.SetHead(0)));
            sheet.addCell(new Label(5, 1, "星期四", scd.SetHead(0)));
            sheet.addCell(new Label(6, 1, "星期五", scd.SetHead(0)));
            sheet.addCell(new Label(7, 1, "星期六", scd.SetHead(0)));
            sheet.addCell(new Label(8, 1, "星期日", scd.SetHead(0)));
            sheet.addCell(new Label(0, 2, "上午", scd.SetHead(1)));
            sheet.addCell(new Label(0, 4, "下午", scd.SetHead(1)));
            sheet.addCell(new Label(0, 6, "晚上", scd.SetHead(1)));
            sheet.addCell(new Label(1, 2, " 8:00 - 9:45", scd.SetHead(2)));
            sheet.addCell(new Label(1, 3, "10:00-11:45", scd.SetHead(2)));
            sheet.addCell(new Label(1, 4, "13:45-15:30", scd.SetHead(2)));
            sheet.addCell(new Label(1, 5, "15:45-17:30", scd.SetHead(2)));
            sheet.addCell(new Label(1, 6, "18:30-20:15", scd.SetHead(2)));
            sheet.addCell(new Label(1, 7, "20:30-22:15", scd.SetHead(2)));
            //合并部分单元格
            sheet.mergeCells(0, 0, 8, 0);
            sheet.mergeCells(0, 2, 0, 3);
            sheet.mergeCells(0, 4, 0, 5);
            sheet.mergeCells(0, 6, 0, 7);
            //设置行高
            sheet.setRowView(0, 600);
            sheet.setRowView(1, 400);
            for (int i = 2; i < 8; i++) {
                sheet.setRowView(i, 1200);
            }
            //设置列宽
            sheet.setColumnView(0, 6);
            sheet.setColumnView(1, 8);
            for (int i = 2; i < 9; i++) {
                sheet.setColumnView(i, 23);
            }
        }
        //创建sheet数组
        WritableSheet[] sheets = wb2.getSheets();
        //创建课程信息的动态数组
        ArrayList<String> Info = new ArrayList<>();

        //正文进行遍历
        for (int row = 2;
                row <= 7; row++) {
            for (int col = 2; col <= 8; col++) {

                //将正文所有表格加上边框
                for (int i = 1; i <= 19; i++) {
                    sheets[i - 1].addCell(new Label(col, row, "", format2));
                }

                //获取总表单元格内容
                String cell = sheet1.getCell(col, row).getContents();
                //定义截取周数的firstposition
                int fpos = -1;
                //定义截取周数的lastposition
                int lpos;
                //定义同一时间冲突课程的第一个课程信息的head位置
                int head = 0;
                //以[作为一个课程的标志查找课程

                while (cell.indexOf("[", fpos + 1) != -1) {
                    //定义每周是否有该课程的数组
                    boolean weeks[] = new boolean[20];
                    for (boolean week : weeks) {
                        week = false;
                    }
                    //排除同样带有[的[考试]的干扰
                    if (cell.indexOf("考试", fpos + 1) != -1) {
                        fpos = cell.indexOf("考试", fpos + 1);
                    }
                    fpos = cell.indexOf("[", fpos + 1);
                    lpos = cell.indexOf("周", fpos + 1);
                    //截取该课程的周数
                    String course = cell.substring(fpos + 1, lpos);
                    //按照，将课程周数截开
                    //例如[1,4,6-8周]操作后得到{"1","4","6-8"}数组
                    String[] wks = course.split("，");
                    //将数组拆成单周并且在weeks[]记录
                    //例如{"1","4","6-8"}操作后得到{1,4,6,7,8}并且对应的5周作标记
                    for (String wk : wks) {
                        int mpos = wk.indexOf("-");
                        if (mpos != -1) {
                            int first = Integer.parseInt(wk.substring(0, mpos));
                            int last = Integer.parseInt(wk.substring(mpos + 1));
                            for (int i = first; i <= last; i++) {
                                weeks[i] = true;
                            }
                        } else {
                            weeks[Integer.parseInt(wk)] = true;
                        }
                    }

                    //获取该课程信息information
                    String info;
                    //课程信息的结束标志是可能出现的冲突课程的标志</br>
                    if (cell.indexOf("<", lpos) != -1) {
                        info = cell.substring(head, cell.indexOf("<", lpos));
                    } else {
                        info = cell.substring(head);
                    }
                    //按照“◇”对格式进行美化
                    StringBuilder sb = new StringBuilder();
                    sb.append(info);
                    int lf = info.indexOf("◇");
                    //注意这里i的作用
                    for (int i = 0; lf != -1; i++) {
                        sb.insert(lf + i, "\n");
                        lf = info.indexOf("◇", lf + 1);
                    }
                    info = sb.toString();

                    //为相同的课程设置颜色
                    boolean isExist = false;
                    int index;
                    WritableCellFormat wcf = new WritableCellFormat(font2);
                    //判断课程信息是否已经存在
                    for (String in : Info) {
                        if (info.compareTo(in) == 0) {
                            isExist = true;
                            index = Info.indexOf(in);
                            wcf = scd.ChangeColor(index);
                        }
                    }
                    if (isExist == false) {
                        Info.add(info);
                        index = Info.indexOf(info);
                        wcf = scd.ChangeColor(index);
                    }

                    //修改对应周数的sheet
                    for (int i = 1; i <= 19; i++) {
                        if (weeks[i] == true) {
                            sheets[i - 1].addCell(new Label(col, row, info, wcf));
                        }
                    }
                    head = cell.indexOf(">", lpos) + 1;
                }
            }
        }
        //写入Excel
        wb2.write();
        //关闭
        wb1.close();
        wb2.close();
    }

    public WritableCellFormat SetHead(int index) throws WriteException {
        //为表头准备颜色数组
        Colour colors[] = {Colour.DARK_BLUE2, Colour.DARK_YELLOW, Colour.YELLOW2,};
        //设置表头格式
        WritableFont wf = new WritableFont(WritableFont.createFont("仿宋"), 12, WritableFont.BOLD);
        WritableCellFormat wcf = new WritableCellFormat(wf);
        wcf.setWrap(true);
        wcf.setAlignment(jxl.format.Alignment.CENTRE);
        wcf.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
        wcf.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THICK, Colour.WHITE);
        wcf.setBackground(colors[index]);
        return wcf;
    }

    public WritableCellFormat ChangeColor(int index) throws WriteException {
        //为课程准备颜色数组
        Colour colors[] = {Colour.TAN, Colour.RED, Colour.PINK, Colour.CORAL, Colour.TEAL,
            Colour.PLUM, Colour.ROSE, Colour.LAVENDER, Colour.GOLD, Colour.BLUE,
            Colour.ORANGE, Colour.LIME, Colour.AQUA, Colour.VIOLET, Colour.INDIGO,
            Colour.GREEN, Colour.YELLOW, Colour.DARK_PURPLE, Colour.DARK_GREEN,
            Colour.GRAY_50, Colour.INDIGO};
        //防止数组越界
        while (index > 20) {
            index -= 20;
        }
        //设置正文格式
        WritableFont wf = new WritableFont(WritableFont.createFont("仿宋"), 11, WritableFont.BOLD);
        wf.setColour(Colour.WHITE);
        WritableCellFormat wcf = new WritableCellFormat(wf);
        wcf.setWrap(true);
        wcf.setAlignment(jxl.format.Alignment.CENTRE);
        wcf.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
        wcf.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.MEDIUM_DASH_DOT_DOT, Colour.WHITE);
        wcf.setBackground(colors[index]);
        return wcf;
    }

}
