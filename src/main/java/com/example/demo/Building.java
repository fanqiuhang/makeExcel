package com.example.demo;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: fanqiuhang
 * Date: 2018/9/3 9:06
 */
public class Building {
    private String buildingNo;//楼号

    private String partNo;//单元号

    private Integer floor; //最高楼层数

    private Integer num; //每层多少户

    public Building(String buildingNo, String partNo, Integer floor, Integer num) {
        this.buildingNo = buildingNo;
        this.partNo = partNo;
        this.floor = floor;
        this.num = num;
    }

    public String getBuildingNo() {
        return buildingNo;
    }

    public void setBuildingNo(String buildingNo) {
        this.buildingNo = buildingNo;
    }

    public String getPartNo() {
        return partNo;
    }

    public void setPartNo(String partNo) {
        this.partNo = partNo;
    }

    public Integer getFloor() {
        return floor;
    }

    public void setFloor(Integer floor) {
        this.floor = floor;
    }

    public Integer getNum() {
        return num;
    }

    public void setNum(Integer num) {
        this.num = num;
    }
}
