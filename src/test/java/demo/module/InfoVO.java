package demo.module;

import cn.afterturn.easypoi.excel.annotation.Excel;

/**
 * 临时车缴费
 */
public class InfoVO {

    @Excel(name = "优惠券编码", width = 15)
    private String couponCode;

    @Excel(name = "第三方用户id", width = 15)
    private String couponUserId;

    @Excel(name = "优惠时间", width = 15)
    private String createTime;

    @Excel(name = "优惠金额", width = 15, numFormat = "0.00")
    private String couponValue;

    public String getCouponCode() {
        return couponCode;
    }

    public void setCouponCode(String couponCode) {
        this.couponCode = couponCode;
    }

    public String getCouponUserId() {
        return couponUserId;
    }

    public void setCouponUserId(String couponUserId) {
        this.couponUserId = couponUserId;
    }

    public String getCreateTime() {
        return createTime;
    }

    public void setCreateTime(String createTime) {
        this.createTime = createTime;
    }

    public String getCouponValue() {
        return couponValue;
    }

    public void setCouponValue(String couponValue) {
        this.couponValue = couponValue;
    }
}
