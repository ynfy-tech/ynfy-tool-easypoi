package demo.module;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.math.BigDecimal;

/**
 * 车辆入场记录
 */
public class ResultExportVO {

	@Excel(name = "停车场", width = 15)
	private String parkName;

    @Excel(name = "支付总金额", width = 15, numFormat = "0.00")
    private BigDecimal payCharge;

    @Excel(name = "总流量", width = 15)
    private Integer totalCount;

    public String getParkName() {
        return parkName;
    }

    public void setParkName(String parkName) {
        this.parkName = parkName;
    }

    public BigDecimal getPayCharge() {
        return payCharge;
    }

    public void setPayCharge(BigDecimal payCharge) {
        this.payCharge = payCharge;
    }

    public Integer getTotalCount() {
        return totalCount;
    }

    public void setTotalCount(Integer totalCount) {
        this.totalCount = totalCount;
    }
}
