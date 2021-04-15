package com.bizcloud.function;

import com.alibaba.fastjson.JSONObject;
import com.bizcloud.ipaas.t97f51c2f74cf4b30bced948079b4d0fd.d20210407113259.auth.extension.AuthConfig;
import com.bizcloud.ipaas.t97f51c2f74cf4b30bced948079b4d0fd.d20210407113259.codegen.TclsqingApi;
import com.bizcloud.ipaas.t97f51c2f74cf4b30bced948079b4d0fd.d20210407113259.model.*;
import com.google.gson.Gson;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.net.URL;
import java.net.URLConnection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HelloFunction {

    public Object handle(Object param, Map<String, String> variables) throws Exception {
        /* 获取ACCESS 权限 */
        AuthConfig authConfig = new AuthConfig(variables.get("APAAS_ACCESS_KEY"), variables.get("APAAS_ACCESS_SECRET"));
        authConfig.initAuth();

        // 传入可选参数调用接口
        HashMap<String, String> options = new HashMap<String, String>();
        options.put("result_type", "json");

        Gson gson = new Gson();

        //差旅申请
        TclsqingApi tclsqingApi = new TclsqingApi();

        String inClass = JSONObject.toJSONString(param);
        JSONObject json = JSONObject.parseObject(inClass);
        //获取t_CLSQing
        String data = JSONObject.toJSONString(json.get("t_CLSQing"));
        JSONObject jsonData = JSONObject.parseObject(data);
        //获取id
        String id = jsonData.get("id").toString();

        //根据id定位查询信息
        TCLSQingDTO query = new TCLSQingDTO();
        query.setId(id);
        //获取该id服务器上存储信息
        List<TCLSQingDTOResponse> list = tclsqingApi.findtCLSQingUsingPOST(query).getData();
        String list_data = gson.toJson(list.get(0));
        JSONObject json_data = JSONObject.parseObject(list_data);

        //获取上传附件的信息
        String scExcel = json_data.get("SCExcel").toString();
        List list_scExcel = JSONObject.parseArray(scExcel);
        //获取附件中的地址
        for (int i = 0; i < list_scExcel.size(); i++) {
            String scExcel_data = gson.toJson(list_scExcel.get(i));
            JSONObject json_scExcel = JSONObject.parseObject(scExcel_data);
            //获取上传Excel 存放地址
            String filePath = json_scExcel.get("filePath").toString();
            System.out.println(filePath);

            URL url = new URL(filePath);
            //调用Excel文档识别
            List<JSONObject> excel_list = readExcel(url);
            for (int j = 0; j < excel_list.size(); j++) {
                String excel_data = gson.toJson(excel_list.get(j));
                JSONObject excel_json = JSONObject.parseObject(excel_data);
                System.out.println(excel_json);
            }
        }
        return null;
    }

    /**
     * Excel文档识别
     *
     * @param url 存放地址
     * @return 返回一个json集合
     * @throws Exception
     */
    public static List<JSONObject> readExcel(URL url) throws Exception {
        // 返回HttpURLConnection 对象
        URLConnection conn = url.openConnection();
        BufferedInputStream bis = null;
        bis = new BufferedInputStream(conn.getInputStream());
        //System.out.println("file type:" + HttpURLConnection.guessContentTypeFromStream(bis));
        //判断格式是否为excel文件
        //url.getPath得到文件的路径
        Workbook workbook = null;
        if (url.getPath().endsWith("xls")) {
            workbook = new HSSFWorkbook(bis);
        } else if (url.getPath().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(bis);
        } else {
            System.out.println("文件格式错误,请检查文件格式！");
        }
        //workbook.getNumberOfSheets() 总共多少页
        //sheet.getPhysicalNumberOfRows() 总共有多少行
        //Excel数据json集合
        List<JSONObject> json_list = new ArrayList<>();
        //每行json数据
        JSONObject json = null;
        for (int k = 0; k < workbook.getNumberOfSheets(); k++) {
            //第一页
            Sheet sheet = workbook.getSheetAt(k);
            //获取内容除去第一行
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                //标题
                List list_title = getTitle(sheet);
                //第i行
                Row row = sheet.getRow(i);
                //第i行第一列
                int start = row.getFirstCellNum();
                //第i行最后一列
                int end = row.getLastCellNum();
                json = new JSONObject();
                String value = "";
                for (int j = start; j < end; j++) {
                    Cell cell = row.getCell(j);
                    //判断是否为空
                    if (row.getCell(j) != null && row.getCell(j).toString() != "") {
                        //根据单元格数值判断
                        //不为空
                        switch (row.getCell(j).getCellType()) {
                            case NUMERIC:   //数值型
                                //判断是否是日期
                                if ("yyyy/mm;@".equals(cell.getCellStyle().getDataFormatString()) || "m/d/yy".equals(cell.getCellStyle().getDataFormatString()) || "yy/m/d".equals(cell.getCellStyle().getDataFormatString()) || "mm/dd/yy".equals(cell.getCellStyle().getDataFormatString()) || "dd-mmm-yy".equals(cell.getCellStyle().getDataFormatString()) || "yyyy/m/d".equals(cell.getCellStyle().getDataFormatString())) {
                                    //值转为日期类型
                                    value = new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                                } else {
                                    //值转为数值类型
                                    value = "" + cell.getNumericCellValue();
                                }
                                //存入json中
                                json.put(list_title.get(j).toString(), value);
                                break;
                            case STRING:    //字符串型
                                //值为String类型
                                value = cell.getStringCellValue();
                                //存入json中
                                json.put(list_title.get(j).toString(), value);
                                break;
                        }
                    } else {
                        //为空
                        value = "";
                        //存入json中
                        json.put(list_title.get(j).toString(), value);
                    }

                }
                //添加到json集合中
                json_list.add(json);
            }
        }
        workbook.close();
        bis.close();
        return json_list;
    }

    /**
     * 获取标题集合
     *
     * @param sheet 默认第一页
     * @return 返回一个标题集合
     */
    public static List getTitle(Sheet sheet) {
        //创建存储标题的集合
        List list_title = new ArrayList();
        //第一行
        Row row = sheet.getRow(0);
        //第一行第一列
        int start = row.getFirstCellNum();
        //第一行最后一列
        int end = row.getLastCellNum();
        for (int i = start; i < end; i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            list_title.add(cell.getStringCellValue());
        }
        return list_title;
    }
}
