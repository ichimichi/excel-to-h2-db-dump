package org.example;

import java.sql.*;

/**
 * Hello world!
 */
public class QueryDemo {
    public static void main(String[] args) {
        try {
            InMemoryDB.getInstance();
            InMemoryDB.dumpExcelToDB("src\\main\\input\\data.xlsx");

//            InMemoryDB.printTable("SELECT * FROM PROVIDER WHERE ACTIVE='Y' AND MTHD='CHK' AND PI='Z'");

            int id = InMemoryDB.getRandomID("SELECT * FROM PROVIDER WHERE ACTIVE='Y'");
//            int id = InMemoryDB.getRandomID("SELECT * FROM PROVIDER WHERE ACTIVE='Y' AND MTHD='CHK' AND PI='Z'");

            InMemoryDB.printTable("SELECT * FROM PROVIDER WHERE ID=" + id);

            ResultSet rs = InMemoryDB.executeQuery("SELECT * FROM PROVIDER WHERE ID=" + id);
            rs.next();

//            System.out.println(rs.getString("FNAME"));
//            System.out.println(rs.getString("LNAME"));
//            System.out.println(rs.getString("PI"));
//            System.out.println(rs.getString("LOC_ID"));
//            System.out.println(rs.getString("MTHD"));
//            System.out.println(rs.getString("ACTIVE"));

            for(int i = 1 ; i <= rs.getMetaData().getColumnCount(); i++){
                System.out.printf("%s:\t%s\n",rs.getMetaData().getColumnName(i), rs.getString(i));
            }
        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            e.printStackTrace();
        } finally {
            InMemoryDB.closeConnection();
        }
    }
}
