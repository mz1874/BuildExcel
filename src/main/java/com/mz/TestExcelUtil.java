package com.mz;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.util.ArrayList;

public class TestExcelUtil {

    private static Logger logger= LoggerFactory.getLogger(TestExcelUtil.class);

    static ArrayList<String> arrayList, stringArrayList;

    static {
        arrayList = new ArrayList<String>(7);
        arrayList.add("GBP");
        arrayList.add("USD");
        arrayList.add("RMB");
        arrayList.add("EUR");
        arrayList.add("KRW");
        arrayList.add("JPY");

        stringArrayList = new ArrayList<>(5);
        stringArrayList.add("ID");
        stringArrayList.add("CAS");
        stringArrayList.add("MDL");
        stringArrayList.add("包装规格描述");
        stringArrayList.add("单位数量");
        stringArrayList.add("货号");
        stringArrayList.add("产品名称");
        stringArrayList.add("分子式");
        stringArrayList.add("纯度");
        stringArrayList.add("包装规格描述");
        stringArrayList.add("单位数量");
        stringArrayList.add("含税总价");
        stringArrayList.add("包装库存状态");
        stringArrayList.add("生产地");
        stringArrayList.add("发货天数");
        stringArrayList.add("失效日期");
        stringArrayList.add("运输条件");
        stringArrayList.add("存储条件");
        stringArrayList.add("状态");
        stringArrayList.add("备注");


    }


    /**
     * 导出excel
     *
     * @throws Exception
     */
    public static void testOutput() throws Exception {
        String filePath = "/";
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("下拉列表测试");
        int lastRowNum = sheet.getLastRowNum();
        String[] datas = new String[]{"维持", "恢复", "调整"};
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
                /*转换时需要指定转换的数组长度*/
                .createExplicitListConstraint(arrayList.toArray(new String[arrayList.size()]));
        CellRangeAddressList addressList = null;
        XSSFDataValidation validation = null;
        /*第0行创建标题*/
        XSSFRow row1 = sheet.createRow(0);
        for (int i = 0; i < stringArrayList.size(); i++) {
            /*设置某一列的宽度*/
            sheet.setColumnWidth(i, 255 * 30);
            XSSFCell cell1 = row1.createCell(i);
            cell1.setCellValue(stringArrayList.get(i));
        } /*设置某一列的宽度*/
        sheet.setColumnWidth(4, 255 * 30);

        for (int i = 1; i < 10; i++) {
            /*创建十行数据*/
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < stringArrayList.size(); j++) {
                XSSFCell cell = row.createCell(j);
                if (j == 4) {
                    cell.setCellValue("请将币种保持一致");
                } else {
                    cell.setCellValue(1);
                }
            }
        }
        addressList = new CellRangeAddressList(0, 9, 4, 4);
        validation = (XSSFDataValidation) dvHelper.createValidation(
                dvConstraint, addressList);
        /**
         * 根据单个询单下的所有商品长度创建
         */

//             07默认setSuppressDropDownArrow(true);
//             validation.setSuppressDropDownArrow(true);
//             validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
        FileOutputStream stream = new FileOutputStream(filePath + "/wd.xlsx");
        workbook.write(stream);
        stream.close();
    }

    public static void main(String[] args) throws Exception {
        testOutput();
    }
}

