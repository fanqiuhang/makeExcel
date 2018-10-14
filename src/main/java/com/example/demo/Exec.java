package com.example.demo;

import java.io.File;
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
        list.add(new Building("4#","A",14,6));
        list.add(new Building("4#","B",14,7));
        list.add(new Building("4#","C",14,7));
        list.add(new Building("5#","A",14,7));
        list.add(new Building("5#","B",14,7));
        list.add(new Building("5#","C",14,7));
        File phone = new File("F://phone.xls");
        Excel.export(list,phone);
    }
}
