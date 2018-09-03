package com.example.demo;

import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: fanqiuhang
 * Date: 2018/9/3 10:59
 */
public class Exec {
    public static void main(String[] args) {
        List list = new ArrayList();
        list.add(new Building("1#","B",14,7));
        list.add(new Building("5#","A",20,8));
        list.add(new Building("3#","C",15,9));
        Excel.export(list);
    }
}
