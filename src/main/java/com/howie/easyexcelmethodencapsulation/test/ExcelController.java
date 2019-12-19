package com.howie.easyexcelmethodencapsulation.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.howie.easyexcelmethodencapsulation.demo.entity.DemoData;
import com.howie.easyexcelmethodencapsulation.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;


/**
 * @author XieShaoping
 */
@RestController
public class ExcelController {
    /**
     * 读取 Excel（允许多个 sheet）
     */
    @RequestMapping(value = "readExcelWithSheets", method = RequestMethod.POST)
    public Object readExcelWithSheets(MultipartFile excel) {
        return ExcelUtil.readExcel(excel, new ImportInfo());
    }

    /**
     * 读取 Excel（指定某个 sheet）
     */
    @RequestMapping(value = "readExcel", method = RequestMethod.POST)
    public Object readExcel(MultipartFile excel, int sheetNo,
                            @RequestParam(defaultValue = "1") int headLineNum) {
        return ExcelUtil.readExcel(excel, new ImportInfo(), sheetNo, headLineNum);
    }

    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        ExcelUtil.writeExcel(response, list, fileName, sheetName, new ExportInfo());
    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";
        ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }

    /**
     * 导出 Excel（一个 sheet,复杂表头）
     */
    @RequestMapping(value = "writeExcelMore", method = RequestMethod.GET)
    public void writeExcelMore(HttpServletResponse response) throws IOException {
        List<MultiLineHeadExcelModel> list = getListMore();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        ExcelUtil.writeExcel(response, list, fileName, sheetName, new MultiLineHeadExcelModel());
    }

    private List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("howie");
        model1.setAge("19");
        model1.setAddress("123456789");
        model1.setEmail("123456789@gmail.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("harry");
        model2.setAge("20");
        model2.setAddress("198752233");
        model2.setEmail("198752233@gmail.com");
        list.add(model2);
        return list;
    }

    private List<MultiLineHeadExcelModel> getListMore() {
        List<MultiLineHeadExcelModel> list = new ArrayList<>();
        MultiLineHeadExcelModel model1 = new MultiLineHeadExcelModel();
        return list;
    }


    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "test", method = RequestMethod.GET)
    public void test(HttpServletResponse response) throws IOException {
        List<DemoData> list = getNewList();
        //String fileName = "一个 Excel 文件";
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), DemoData.class).sheet("模板").doWrite(list);
    }

    private List<DemoData> getNewList() {
        List<DemoData> list = new ArrayList<>();
        DemoData model1 = new DemoData();
        model1.setDate(new Date());
        model1.setDoubleData(19.00);
        model1.setString("123456789");
        list.add(model1);
        DemoData model2 = new DemoData();
        model2.setDate(new Date());
        model2.setDoubleData(19.00);
        model2.setString("123456789");
        list.add(model2);
        return list;
    }


    /**
     * 导出 Excel,多个 sheet,同一个对象,如果写到不同的sheet 同一个对象
     */
    @GetMapping("testMoreByOneVo")
    public void testMore(HttpServletResponse response) throws IOException {
        //头部单元样式
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        //内容单元样式
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        //内容居中
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //水平单元格样式策略
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        String fileName = "test" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = ExcelUtil.write(fileName, response, DemoData.class);
        // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
        for (int i = 0; i < 5; i++) {
            // 每次都要创建writeSheet 这里注意必须指定sheetNo,sheetName不能重复
            WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i)
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .build();
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<DemoData> data = getNewList();
            excelWriter.write(data, writeSheet);
        }
        // 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }
}
