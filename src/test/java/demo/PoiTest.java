package demo;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.google.common.collect.Lists;
import demo.config.DemoExcelStyleBean;
import demo.module.ExportVO;
import demo.module.InfoVO;
import demo.module.ResultExportVO;
import org.junit.Test;
import tech.ynfy.tool.easypoi.module.PoiBO;
import tech.ynfy.tool.easypoi.util.PoiUtil;

import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 〈〉
 *
 * @author Hsiong
 * @version 1.0.0
 * @since 2022/8/8
 */
public class PoiTest {

    @Test
    public void poiTest() {

        // sheet result
        List<ResultExportVO> carResultList = Lists.newArrayList();

        Map<String, List<ExportVO>> exportSheetVOS = new HashMap<>();
        // sheet 1
        List<ExportVO> sheetList1 = Lists.newArrayList();
        // sheet 1 row 1
        ExportVO sheetVO1 = new ExportVO();
        sheetVO1.setPlate("car 1");
        sheetVO1.setParkName("park 1");
        sheetVO1.setInTime("2022-01-01");
        sheetVO1.setOutTime("2022-01-02");
        sheetVO1.setUserId("id 1");
        sheetVO1.setPayTime(new Date());
        sheetVO1.setPayCharge(new BigDecimal(5));
        sheetVO1.setOutTradeNo("order 1");
        // sheet 1 row 1 coupon row 1
        List<InfoVO> couponInfo = Lists.newArrayList();
        InfoVO infoVO1 = new InfoVO();
        infoVO1.setCouponCode("code 1");
        infoVO1.setCouponUserId("out id 1");
        infoVO1.setCreateTime("2022-01-01");
        infoVO1.setCouponValue("3");
        // sheet 1 row 1 coupon row 2
        InfoVO infoVO2 = new InfoVO();
        infoVO2.setCouponCode("code 2");
        infoVO2.setCouponUserId("out id 2");
        infoVO2.setCreateTime("2022-01-02");
        infoVO2.setCouponValue("2");
        couponInfo.add(infoVO1);
        couponInfo.add(infoVO2);
        sheetVO1.setCouponInfo(couponInfo);

        ResultExportVO sheetResult1 = new ResultExportVO();
        sheetResult1.setParkName(sheetVO1.getParkName());
        sheetResult1.setPayCharge(sheetVO1.getPayCharge());
        sheetResult1.setTotalCount(1);

        // sheet 1 row 2
        ExportVO sheetVO2 = new ExportVO();
        sheetVO2.setPlate("car 2");
        sheetVO2.setParkName("park 2");
        sheetVO2.setInTime("2022-02-02");
        sheetVO2.setOutTime("2022-02-02");
        sheetVO2.setUserId("id 2");
        sheetVO2.setPayTime(new Date());
        sheetVO2.setPayCharge(new BigDecimal(5));
        sheetVO2.setOutTradeNo("order 2");
        List<InfoVO> couponInfo2 = Lists.newArrayList();
        couponInfo2.add(infoVO1);
        sheetVO2.setCouponInfo(couponInfo2);
        sheetList1.add(sheetVO1);
        sheetList1.add(sheetVO2);
        exportSheetVOS.put("sheet 1", sheetList1);

        ResultExportVO sheetResult2 = new ResultExportVO();
        sheetResult2.setParkName(sheetVO1.getParkName());
        sheetResult2.setPayCharge(sheetVO1.getPayCharge());
        sheetResult2.setTotalCount(1);


        // sheet 2
        List<ExportVO> sheetList2 = Lists.newArrayList();
        sheetList2.add(sheetVO2);
        exportSheetVOS.put("sheet 2", sheetList2);

        // count
        sheetResult2.setPayCharge(sheetResult2.getPayCharge().add(sheetVO2.getPayCharge()));
        sheetResult2.setTotalCount(sheetResult2.getTotalCount() + 1);
        carResultList.add(sheetResult1);
        carResultList.add(sheetResult2);

        this.exportCarInXls(exportSheetVOS, carResultList);

    }

    /**
     * 导出 service demo
     *
     * @param exportSheetVOS 各工作表数据, key,value key: sheetName, value: data
     * @param carResultList  结果 sheet 页
     */
    private void exportCarInXls(Map<String, List<ExportVO>> exportSheetVOS, List<ResultExportVO> carResultList) {

        String fileName = "导出测试统计(2022.08.05-2022.08.10)";
        // 默认工作表名
        String initSheetName = "导出结果";

        List<PoiBO> demoPoiBOS = Lists.newArrayList();

        /**
         * 各 sheet 页
         */
        for (Map.Entry<String, List<ExportVO>> entry : exportSheetVOS.entrySet()) {
            PoiBO poiBO = new PoiBO();
            poiBO.setSheetName(entry.getKey());
            poiBO.setData(entry.getValue());
            // 有 groupName, 工作表会从第三列开始
            poiBO.setStartRow(3);
            demoPoiBOS.add(poiBO);
        }

        /**
         * 结果 sheet 页
         */
        PoiBO poiBO2 = new PoiBO();
        poiBO2.setSheetName("合计");
        poiBO2.setData(carResultList);
        poiBO2.setClassOrder(1);
        poiBO2.setStartRow(2);
        demoPoiBOS.add(poiBO2);

        /**
         * 导出
         */
        ExportParams exportParams = this.getExportParams(fileName, initSheetName);
        Class<?>[] pojoClassList = new Class[]{ExportVO.class, ResultExportVO.class};
        String fileDir = PoiUtil.getInstance().exportBigDataXls(fileName, exportParams, demoPoiBOS, pojoClassList);
        System.out.println("fileDir " + fileDir);

    }

    /**
     * 导出参数
     *
     * @param name
     * @return
     */
    private ExportParams getExportParams(String name, String initSheetName) {
        //表格名称,sheet名称,导出版本
        ExportParams exportParams = new ExportParams(name, initSheetName, ExcelType.XSSF);
        exportParams.setStyle(DemoExcelStyleBean.class);
        return exportParams;
    }


}
