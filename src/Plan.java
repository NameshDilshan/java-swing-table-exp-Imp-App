/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Batman
 */
public class Plan {
    private String OrderNo;
    private String ItemCode;
    private String QTY;
    private String Completed_QTY;
    private String ODate;
    private String OrderedBy;
    private String Color;
    private String WEIGHT;
    private String Completed_WEIGHT;

    public String getOrderNo() {
        return OrderNo;
    }

    public void setOrderNo(String OrderNo) {
        this.OrderNo = OrderNo;
    }

    public String getItemCode() {
        return ItemCode;
    }

    public void setItemCode(String ItemCode) {
        this.ItemCode = ItemCode;
    }

    public String getQTY() {
        return QTY;
    }

    public void setQTY(String QTY) {
        this.QTY = QTY;
    }

    public String getCompleted_QTY() {
        return Completed_QTY;
    }

    public void setCompleted_QTY(String Completed_QTY) {
        this.Completed_QTY = Completed_QTY;
    }

    public String getODate() {
        return ODate;
    }

    public void setODate(String ODate) {
        this.ODate = ODate;
    }

    public String getOrderedBy() {
        return OrderedBy;
    }

    public void setOrderedBy(String OrderedBy) {
        this.OrderedBy = OrderedBy;
    }

    public String getColor() {
        return Color;
    }

    public void setColor(String Color) {
        this.Color = Color;
    }

    public String getWEIGHT() {
        return WEIGHT;
    }

    public void setWEIGHT(String WEIGHT) {
        this.WEIGHT = WEIGHT;
    }

    public String getCompleted_WEIGHT() {
        return Completed_WEIGHT;
    }

    public void setCompleted_WEIGHT(String Completed_WEIGHT) {
        this.Completed_WEIGHT = Completed_WEIGHT;
    }

    public Plan(String OrderNo, String ItemCode, String QTY, String Completed_QTY, String ODate, String OrderedBy, String Color, String WEIGHT, String Completed_WEIGHT) {
        this.OrderNo = OrderNo;
        this.ItemCode = ItemCode;
        this.QTY = QTY;
        this.Completed_QTY = Completed_QTY;
        this.ODate = ODate;
        this.OrderedBy = OrderedBy;
        this.Color = Color;
        this.WEIGHT = WEIGHT;
        this.Completed_WEIGHT = Completed_WEIGHT;
    }
    
    
    
    
}
