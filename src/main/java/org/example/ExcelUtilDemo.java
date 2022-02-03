package org.example;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Random;

public class ExcelUtilDemo {
    public static void main(String[] args) {
        ExcelUtil excelUtil = new ExcelUtil("src\\main\\input\\data.xlsx");

        excelUtil.printStats();
//        excelUtil.printSheet("PROVIDER");
        System.out.println(excelUtil.getCellData("CLAIM", 0, 0));
//        excelUtil.setCellData("CLAIM", 0, "DATE", "HELLO");
//        excelUtil.setCellData("CLAIM", 1, "DATE", "HELLO");
        int row = excelUtil.getRowCount("CLAIM");
        LocalDate date = LocalDate.now();
        Random rand = new Random();

        String format1 = "yyMMdd";

        DateTimeFormatter dtf1 = DateTimeFormatter.ofPattern(format1);

        for(int i = 0; i < 30 ; i++){
            String fcnDate = date.format(dtf1);
            StringBuilder fcnSeq = new StringBuilder();
            fcnSeq.append(String.valueOf(rand.nextInt(100000)));

            while(fcnSeq.length() < 5){
                fcnSeq.insert(0, "0");
            }

            String officeCode = String.valueOf(rand.nextInt(1000));
            String bulk = rand.nextInt(2) > 0 ? "Y" : "N";
            String pass = rand.nextInt(2) > 0 ? "PASS" : "FAIL";

            System.out.println(fcnDate);
            System.out.println(fcnSeq);
            System.out.println(fcnDate + fcnSeq);
            System.out.println(officeCode);
            System.out.println(bulk);
            System.out.println(pass);

            excelUtil.setCellData("CLAIM", row + i, "DATE", "22" + fcnDate);
            excelUtil.setCellData("CLAIM", row + i, "SEQ", fcnSeq.toString());
            excelUtil.setCellData("CLAIM", row + i, "FCN", fcnDate + fcnSeq);
            excelUtil.setCellData("CLAIM", row + i, "OFFICE_CODE", officeCode);
            excelUtil.setCellData("CLAIM", row + i, "BULK", bulk);
            excelUtil.setCellData("CLAIM", row + i, "PASS", pass);
        }


    }
}
