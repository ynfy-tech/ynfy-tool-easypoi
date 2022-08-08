package demo.module;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import com.google.common.collect.Lists;
import org.springframework.format.annotation.DateTimeFormat;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

/**
 * 车辆入场记录
 */
public class ExportVO {

    @Excel(name = "车牌", width = 15)
    private String plate;

	@Excel(name = "停车场", width = 15, groupName = "停车场信息")
	private String parkName;

	@Excel(name = "入场时间", width = 15, groupName = "停车场信息")
	private String inTime;

    @Excel(name = "出场时间", width = 15, groupName = "停车场信息")
    private String outTime;

    @Excel(name = "用户id", width = 15, groupName = "缴费信息")
    private String userId;

    @Excel(name = "缴费时间", width = 20, format = "yyyy-MM-dd HH:mm:ss", groupName = "缴费信息")
    @DateTimeFormat(pattern = "yyyy-MM-dd HH:mm:ss")
    private Date payTime;

    @Excel(name = "缴费金额", width = 15, groupName = "缴费信息", numFormat = "0.00")
    private java.math.BigDecimal payCharge;

    @Excel(name = "支付单号", width = 15, groupName = "缴费信息")
    private String outTradeNo;

    @ExcelCollection(name = "优惠信息", orderNum = "99")
    List<InfoVO> couponInfo = Lists.newArrayList();

    public String getPlate() {
        return plate;
    }

    public void setPlate(String plate) {
        this.plate = plate;
    }

    public String getParkName() {
        return parkName;
    }

    public void setParkName(String parkName) {
        this.parkName = parkName;
    }

    public String getInTime() {
        return inTime;
    }

    public void setInTime(String inTime) {
        this.inTime = inTime;
    }

    public String getOutTime() {
        return outTime;
    }

    public void setOutTime(String outTime) {
        this.outTime = outTime;
    }

    public String getUserId() {
        return userId;
    }

    public void setUserId(String userId) {
        this.userId = userId;
    }

    public Date getPayTime() {
        return payTime;
    }

    public void setPayTime(Date payTime) {
        this.payTime = payTime;
    }

    public BigDecimal getPayCharge() {
        return payCharge;
    }

    public void setPayCharge(BigDecimal payCharge) {
        this.payCharge = payCharge;
    }

    public String getOutTradeNo() {
        return outTradeNo;
    }

    public void setOutTradeNo(String outTradeNo) {
        this.outTradeNo = outTradeNo;
    }

    public List<InfoVO> getCouponInfo() {
        return couponInfo;
    }

    public void setCouponInfo(List<InfoVO> couponInfo) {
        this.couponInfo = couponInfo;
    }
}
