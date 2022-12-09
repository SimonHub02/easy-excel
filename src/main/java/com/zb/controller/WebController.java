package com.zb.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.zb.pojo.User;
import com.zb.write.style.ExcelCellWidthStyleStrategy;
import com.zb.read.ExcelMergeHelper;
import com.zb.write.style.CellStyleStrategy;
import com.zb.write.ExcelFillCellMergeStrategy;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
public class WebController {


    /**
     * 行合并的读写方式
     *
     * @param file excel文档
     * @return 数据集
     * @throws Exception 异常
     */
    @PostMapping("/read")
    public List<User> show(@RequestParam("file") MultipartFile file) throws Exception {
        ExcelMergeHelper<User> helper = new ExcelMergeHelper();
        List<User> list = helper.getList(file, User.class, 0, 1);
        return list;
    }

    /**
     * 单sheet导出
     *
     * @param response 响应
     */
    @PostMapping("write")
    public void show(HttpServletResponse response) {
        try {
            //来点假数据
            List<User> list = new ArrayList<>();
            Date date = new Date();
            for (int i = 0; i < 5; i++) {
                User user = new User();
                user.setName("test" + i);
                user.setAge("1" + i);
                user.setNo("2" + 1);
                user.setDate(date);
                list.add(user);
            }

            // 设置第几列合并
            //这边我需要指定合并第一列，所以赋值0
            // 如果需要合并多列，直接逗号分隔：int[] mergeColumnIndex = {0,1,2}
            int[] mergeColumnIndex = {2};
            // 需要从第几行开始合并
            int mergeRowIndex = 1;
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            //  设置文件名称
            String fileName = URLEncoder.encode("测试导出", "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            //  sheet名称
            EasyExcel.write(response.getOutputStream(),
                            User.class) //用户
                    //样式剧中
                    .registerWriteHandler(horizontalCellStyleStrategyBuilder())
                    //excel版本
                    .excelType(ExcelTypeEnum.XLSX)
                    //是否自动关流
                    .autoCloseStream(Boolean.TRUE)
                    .registerWriteHandler(new ExcelFillCellMergeStrategy(mergeRowIndex, mergeColumnIndex)).
                    sheet("测试导出").doWrite(list);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 多sheet导出
     *
     * @param response 响应
     */
    @PostMapping("/writeB")
    public void write(HttpServletResponse response) {
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            //  设置文件名称
            String fileName = URLEncoder.encode("test", "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            List<User> temp = new ArrayList<>();
            Date date = new Date();
            for (int i = 0; i < 5; i++) {
                User user = new User();
                user.setName("test" + i);
                user.setAge("1" + i);
                user.setNo("2" + 1);
                user.setDate(date);
                temp.add(user);
            }
            int[] mergeColumnIndex = {2};
            // 需要从第几行开始合并
            int mergeRowIndex = 1;
            OutputStream stream = response.getOutputStream();
            ExcelWriter excelWriter = EasyExcel.write(stream)
                    .registerWriteHandler(horizontalCellStyleStrategyBuilder())
                    .registerWriteHandler(new ExcelCellWidthStyleStrategy())
                    .registerWriteHandler(new ExcelFillCellMergeStrategy(mergeRowIndex, mergeColumnIndex))
                    .excelType(ExcelTypeEnum.XLSX)
                    .autoCloseStream(true)
                    .build();
            WriteSheet sheet = EasyExcel.writerSheet(0, "test1").head(User.class).build();
            excelWriter.write(temp, sheet);
            WriteSheet sheet1 = EasyExcel.writerSheet(1, "test2").head(User.class).build();
            excelWriter.write(temp, sheet1);
            excelWriter.finish();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 单元格格式配置
     *
     * @return
     */
    public CellStyleStrategy horizontalCellStyleStrategyBuilder() {
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        //设置头字体
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 13);
        headWriteFont.setBold(true);
        headWriteCellStyle.setWriteFont(headWriteFont);
        //设置头居中
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //内容策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        //设置 水平居中
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return new CellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }
}
