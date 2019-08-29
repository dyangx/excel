package com.example.demo.util;

import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

public class ExcelUtil {


    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> clazz,String fileName,
                                   HttpServletResponse response) throws IOException {
        exportExcel(list,title,sheetName,clazz,fileName,response,null);
    }
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> clazz,String fileName,
                                   HttpServletResponse response,List<String> columns) throws IOException {
        //表格样式构造
        ExportParams exportParams = new ExportParams(title, sheetName);
        Workbook workbook = getWorkbook(exportParams,clazz,list,columns);
        downLoad(fileName, response, workbook);
    }

    public static Workbook getWorkbook(ExportParams entity, Class<?> clazz, Collection<?> dataSet,List<String> columns) {
        //创建一个workbook
        Workbook workbook = createWorkbook(entity.getType(), dataSet.size());
        //将数据写入workbook
        writeToWookbook(workbook, entity, clazz, dataSet,columns);
        return workbook;
    }

    public static void writeToWookbook(Workbook workbook, ExportParams entity, Class<?> clazz, Collection<?> dataSet,List<String> columns) {
        ExcelExportService excelExportService = new ExcelExportService();
        if (workbook != null && entity != null && clazz != null && dataSet != null) {
            try {
                List<ExcelExportEntity> excelParams = new ArrayList();
                Field[] fileds = PoiPublicUtil.getClassFields(clazz);
                ExcelTarget etarget = clazz.getAnnotation(ExcelTarget.class);
                String targetId = etarget == null ? null : etarget.value();
                excelExportService.getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, clazz, null, null);
                removeParams(excelParams,columns);
                excelExportService.createSheetForMap(workbook, entity, excelParams, dataSet);
            } catch (Exception var9) {
                throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, var9.getCause());
            }
        } else {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
    }

    /**
     * @description
     */
    private static void removeParams(List<ExcelExportEntity> params,List<String> columns){
        if(columns == null || columns.size() == 0) return;
        Iterator<ExcelExportEntity> iterator = params.iterator();
        while (iterator.hasNext()){
            ExcelExportEntity entity = iterator.next();
            for (String column : columns){
                if(Objects.equals(column,entity.getName())){
                    iterator.remove();
                    break;
                }
            }
        }
    }

    /**
     * @description 创建一个空白Wookbook
     */
    private static Workbook createWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else {
            return size < 100000 ? new XSSFWorkbook() : new SXSSFWorkbook();
        }
    }

    /**
     * @description 下载
     */
    public static void downLoad(String fileName, HttpServletResponse response, Workbook workbook) throws IOException {
        response.setCharacterEncoding("UTF-8");
        response.setHeader("content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition",
                "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        workbook.write(response.getOutputStream());
    }
}
