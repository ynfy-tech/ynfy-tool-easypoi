package tech.ynfy.tool.easypoi.util;

import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.excel.export.styler.IExcelExportStyler;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import cn.afterturn.easypoi.util.PoiExcelGraphDataUtil;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 提供批次插入服务
 */
public class ExcelBatchExportUtil extends ExcelExportService {

    private static ThreadLocal<ExcelBatchExportUtil> THREAD_LOCAL = new ThreadLocal<ExcelBatchExportUtil>();

    private Workbook                                   workbook;
    private Sheet                                      sheet;
    private List<ExcelExportEntity>                    excelParams;
    private ExportParams                               entity;
    private int                                        titleHeight;
    private Drawing                                    patriarch;
    private short                                      rowHeight;
    private int                                        index;

    @Override
    public void insertDataToSheet(Workbook workbook,
                                  ExportParams entity,
                                  List<ExcelExportEntity> entityList,
                                  Collection<?> dataSet,
                                  Sheet sheet) {
        try {
            dataHandler = entity.getDataHandler();
            if (dataHandler != null && dataHandler.getNeedHandlerFields() != null) {
                needHandlerList = Arrays.asList(dataHandler.getNeedHandlerFields());
            }
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) entity.getStyle()
                .getConstructor(Workbook.class).newInstance(workbook));
            patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            this.index = entity.isCreateHeadRows()
                ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            titleHeight = index;
            setCellWith(excelParams, sheet);
            rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    /**
     * 单例模式初始化工具类
     * @param entity 报表参数
     * @param pojoClass 报表class
     * @return 返回该工具类实例
     */
    protected static ExcelBatchExportUtil getExcelBatchExportService(ExportParams entity,
                                                                     Class<?> pojoClass) {
        if (THREAD_LOCAL.get() == null) {
            ExcelBatchExportUtil batchServer = new ExcelBatchExportUtil();
            batchServer.init(entity, pojoClass);
            THREAD_LOCAL.set(batchServer);
        }
        return THREAD_LOCAL.get();
    }

    /**
     * 初始化
     * @param entity 报表参数
     * @param pojoClass 报表class
     */
    protected void init(ExportParams entity, Class<?> pojoClass) {
        List<ExcelExportEntity> excelParams = createExcelExportEntityList(entity, pojoClass);
        init(entity, excelParams);
    }

    /**
     * 初始化
     * @param entity 报表参数
     * @param excelParams excel 导出工具类,对cell类型做映射
     */
    protected void init(ExportParams entity, List<ExcelExportEntity> excelParams) {
        LOGGER.debug("ExcelBatchExportServer only support SXSSFWorkbook");
        entity.setType(ExcelType.XSSF);
        workbook = new SXSSFWorkbook();
        this.entity = entity;
        this.excelParams = excelParams;
        super.type = entity.getType();
        createSheet(workbook, entity, excelParams);
        if (entity.getMaxNum() == 0) {
            entity.setMaxNum(1000000);
        }
        insertDataToSheet(workbook, entity, excelParams, null, sheet);
    }

    /**
     * 初始化报表实体
     * @param entity 报表参数
     * @param pojoClass 报表class
     * @return
     */
    protected List<ExcelExportEntity> createExcelExportEntityList(ExportParams entity, Class<?> pojoClass) {
        try {
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            // 得到所有字段
            Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
                             null, null);
            sortAllParams(excelParams);

            return excelParams;
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    /**
     * 初始化报表
     * @param workbook 报表
     * @param entity 导出实体
     * @param excelParams 导出类参数
     */
    protected void createSheet(Workbook workbook, ExportParams entity, List<ExcelExportEntity> excelParams) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel export start ,List<ExcelExportEntity> is {}", excelParams);
            LOGGER.debug("Excel version is {}",
                         entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || excelParams == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        try {
            try {
                sheet = workbook.createSheet(entity.getSheetName());
            } catch (Exception e) {
                // 重复遍历,出现了重名现象,创建非指定的名称Sheet
                sheet = workbook.createSheet();
            }
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    /**
     * 给当前工作表添加数据
     * @param dataSet 数据
     * @param ifNewPage 是否新建新的工作表
     * @param newStartRow 新建行数
     * @param newSheet 新建工作表
     * @param newExcelParams 导出参数
     * @return 返回工作表
     */
    protected Workbook appendData(Collection<?> dataSet, boolean ifNewPage, int newStartRow, Sheet newSheet, List<ExcelExportEntity> newExcelParams) {
        if (ifNewPage) {
            index = newStartRow;
            sheet = newSheet;
            if (PoiUtil.checkArray(newExcelParams)) {
                excelParams = newExcelParams;
            }
        }

        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            try {
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
            } catch (Exception e) {
                LOGGER.error(e.getMessage(), e);
                throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
            }
        }
        return workbook;
    }

/*    protected static ExcelBatchExportUtil getExcelBatchExportService(ExportParams entity,
                                                                  List<ExcelExportEntity> excelParams) {
        if (THREAD_LOCAL.get() == null) {
            ExcelBatchExportUtil batchServer = new ExcelBatchExportUtil();
            batchServer.init(entity, excelParams);
            THREAD_LOCAL.set(batchServer);
        }
        return THREAD_LOCAL.get();
    }

    protected static ExcelBatchExportUtil getCurrentExcelBatchExportService() {
        return THREAD_LOCAL.get();
    }

    protected void closeExportBigExcel() {
        if (entity.getFreezeCol() != 0) {
            sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
        }
        mergeCells(sheet, excelParams, titleHeight);
        // 创建合计信息
        addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
        THREAD_LOCAL.remove();

    }*/

}
