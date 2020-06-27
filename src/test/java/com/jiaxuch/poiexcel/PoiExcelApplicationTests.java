package com.jiaxuch.poiexcel;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class PoiExcelApplicationTests {

    @Test
    void contextLoads() {
        String num = "1.0";
        long l = Long.parseLong(num);
        System.out.println(l);


    }


}
