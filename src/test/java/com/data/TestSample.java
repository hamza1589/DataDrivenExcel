package com.data;

import java.io.IOException;
import java.util.ArrayList;

public class TestSample {
    public static void main(String[] args) throws IOException {
        DataDriven d= new DataDriven();
        ArrayList data=d.getData("add profile");
        System.out.println(data.get(0));
        System.out.println(data.get(1));
        System.out.println(data.get(2));
        System.out.println(data.get(3));
    }
}
