package tech.ynfy.tool.easypoi.util;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tech.ynfy.tool.easypoi.module.PoiBO;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.util.List;
import java.util.Map;

/**
 * 增强型 POI 工具类
 */
public class PoiUtil {

    private static class Inner {
        private static final PoiUtil INSTANCE = new PoiUtil();
    }

    /***
     * 单例模式之：静态内部类单例模式
     * 只有第一次调用getInstance方法时，虚拟机才加载 Inner 并初始化instance ，只有一个线程可以获得对象的初始化锁，其他线程无法进行初始化，
     * 保证对象的唯一性。目前此方式是所有单例模式中最推荐的模式，但具体还是根据项目选择。
     * @return
     */
    public static PoiUtil getInstance() {
        return Inner.INSTANCE;
    }

    /**
     * 多表单 导出大数据
     *
     * @param fileName      文件名
     * @param params        表格样式
     * @param poiBOList     表单数据
     * @param pojoClassList 表单数据类型
     * @return 导出excel位置
     */
    public String exportBigDataXls(String fileName,
                                   ExportParams params,
                                   List<PoiBO> poiBOList,
                                   Class<?>... pojoClassList) {
        String fileDir = new File("").getAbsolutePath() + File.separator + fileName + ".xlsx";
        File file = new File(fileDir);
        
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                throw new IllegalArgumentException(e.getMessage(), e);
            }
        }

        Class<?> pojoClass = pojoClassList[0];

        ExcelBatchExportUtil batchExportUtil = ExcelBatchExportUtil.getExcelBatchExportService(params, pojoClass);
        List<ExcelExportEntity> excelExportEntities = batchExportUtil.createExcelExportEntityList(params, pojoClass);

        Workbook workbook = null;
        for (PoiBO poiBO : poiBOList) {
            String sheetName = poiBO.getSheetName();
            List<?> collection = poiBO.getData();
            Integer startRow = poiBO.getStartRow();

            boolean ifNewpage = false;
            Sheet sheet = null;
            List<ExcelExportEntity> newExcelParams = null;
            if (checkArray(workbook)) {
                Object o = poiBO.getClassOrder();
                if (checkArray(o)) {
                    Integer order = (Integer) o;
                    pojoClass = pojoClassList[order];
                    excelExportEntities = batchExportUtil.createExcelExportEntityList(params, pojoClass);
                    newExcelParams = excelExportEntities;
                }

                sheet = workbook.createSheet(sheetName);
                batchExportUtil.insertDataToSheet(workbook, params, excelExportEntities, null, sheet);
                ifNewpage = true;
            }

            workbook = batchExportUtil.appendData(collection, ifNewpage, startRow, sheet, newExcelParams);

        }

        try (BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(fileDir));) {
            workbook.write(outputStream);
            outputStream.flush();
            // 释放workbook所占用的所有资源
            workbook.close();
        } catch (IOException e) {
            throw new IllegalArgumentException(e.getMessage(), e);
        }
        return fileDir;
    }

    /**
     * 利用反射 改变order排序
     *
     * @param exportVOS 导出结果
     * @param <T>       泛型
     */
    public <T> void changeExportOrder(List<T> exportVOS) {
        /**
         * 批量添加groupname排序
         */
        if (checkArray(exportVOS)) {
            try {
                T exportVO = exportVOS.get(0);
                Field[] fields = exportVO.getClass().getDeclaredFields();
                int count = 0;
                for (int i = 0; i < fields.length; i++) {
                    Field field = fields[i];
                    // 修改排序
                    Excel excel = field.getAnnotation(Excel.class);
                    if (checkArray(excel)) {
                        if (StringUtils.isNotEmpty(excel.groupName())) {
                            InvocationHandler invocationHandler = Proxy.getInvocationHandler(excel);
                            // 获取 AnnotationInvocationHandler 的 memberValues 字段
                            Field declaredField = invocationHandler.getClass().getDeclaredField("memberValues");
                            // 因为这个字段事 private final 修饰，所以要打开权限
                            declaredField.setAccessible(true);
                            // 获取 memberValues
                            Map memberValues = (Map) declaredField.get(invocationHandler);
                            // 修改 value 属性值
                            memberValues.put("orderNum", count + "");
                            // 获取 foo 的 value 属性值
                            // String newValue = excel.orderNum();
                            // System.out.println("修改之后的注解值：" + newValue);
                        }
                    }
                    count++;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 检查是否为空
     *
     * @param obj 对象
     * @return 返回不为空true ; 为空则为 false
     */
    protected static boolean checkArray(Object obj) {
        if (null == obj) {
            return false;
        }

        if (obj instanceof List) {
            return ((List) obj).size() > 0;
        }

        return true;
    }
}
