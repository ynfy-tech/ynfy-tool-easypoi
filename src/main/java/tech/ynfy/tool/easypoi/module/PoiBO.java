package tech.ynfy.tool.easypoi.module;


import java.util.List;

/**
 * demoBO
 */
public class PoiBO {

    /**
     * sheet名
     */
    private String sheetName;

    /**
     * 数据
     */
    private List<?> data;

    /**
     * 开始填充行
     */
    private Integer startRow;

    /************************************  选填 从0开始  ********************************************/

    /**
     * 取第几个类
     */
    private Integer classOrder;


    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getClassOrder() {
        return classOrder;
    }

    public void setClassOrder(Integer classOrder) {
        this.classOrder = classOrder;
    }
}
