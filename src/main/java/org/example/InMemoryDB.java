package org.example;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Random;

public class InMemoryDB {
    private static final String DB_DRIVER = "org.h2.Driver";
    private static final String DB_CONNECTION = "jdbc:h2:mem:test";
    private static final String DB_USER = "";
    private static final String DB_PASSWORD = "";
    private static Connection conn = null;
    private static InMemoryDB instance;
    public static HashMap<String, ArrayList<String>> tableFieldMap;

    private InMemoryDB() throws SQLException, ClassNotFoundException {
        getConnection();
    }

    public static InMemoryDB getInstance() throws SQLException, ClassNotFoundException {
        if (instance == null) {
            instance = new InMemoryDB();
        }

        return instance;
    }

    private static Connection getConnection() throws ClassNotFoundException, SQLException {
        conn = null;
        Class.forName(DB_DRIVER);
        conn = DriverManager.getConnection(DB_CONNECTION, DB_USER, DB_PASSWORD);
        System.out.println(conn);
        return conn;

    }

    public static void closeConnection() {
        try {
            if (conn != null) {
                conn.close();
            }

            if (instance != null) {
                instance = null;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static ResultSet executeQuery(String query) throws SQLException {
        ResultSet rs;
        rs = conn.createStatement().executeQuery(query);
        System.out.println(rs);
        return rs;
    }

    private static int getRandomIndex(int size){
        Random rand = new Random();
        return rand.nextInt(size);
    }

    public static int getRandomID(String query) throws SQLException {
        ResultSet rs = executeQuery(query);
        ArrayList<Integer> idList = new ArrayList<>();
        int id;

        while (rs.next()) {
            idList.add(rs.getInt("ID"));
        }

        try{
            id = idList.get(getRandomIndex(idList.size()));
        }catch (Exception e){
            id = Integer.MIN_VALUE;
        }

        return id;
    }

    public static void printTable(String query) throws SQLException {
        ResultSet rs = executeQuery(query);

        for (int i = 1; i <= rs.getMetaData().getColumnCount(); i++) {
            System.out.printf("%s\t\t", rs.getMetaData().getColumnName(i));
        }
        System.out.println();

        StringBuilder line = new StringBuilder();
        for (int i = 0; i < 10 * rs.getMetaData().getColumnCount(); i++) {
            line.append("-");
        }

        System.out.println(line);

        while (rs.next()) {
            for (int i = 1; i <= rs.getMetaData().getColumnCount(); i++) {
                System.out.printf("%s\t\t", rs.getString(i));
            }
            System.out.println();
        }
    }

    public static void dumpExcelToDB(String path) throws SQLException {
        ExcelUtil excelUtil = new ExcelUtil(path);
        tableFieldMap = new HashMap<>();

        for (String sheet : excelUtil.getSheetNames()) {
            int rowNum = excelUtil.getRowCount(sheet);
            int colNum = excelUtil.getColumnCount(sheet);


            ArrayList<String> fieldList = new ArrayList<>();
            ArrayList<String> fieldTypeList = new ArrayList<>();

            fieldList.add("ID");
            fieldTypeList.add("IDENTITY NOT NULL PRIMARY KEY");

            for (int j = 0; j < colNum; j++) {
                fieldList.add(excelUtil.getCellData(sheet, 0, j));
            }
            tableFieldMap.put(sheet, fieldList);

            for (int j = 0; j < colNum; j++) {
//                fieldTypeList.add(excelUtil.getCellData(sheet, 1, j));
                fieldTypeList.add("VARCHAR(255)");
            }

            StringBuilder createQuery = new StringBuilder();
            createQuery.append("CREATE TABLE ");
            createQuery.append(sheet);
            createQuery.append("(");
            for (int i = 0; i < fieldList.size(); i++) {
                createQuery.append(fieldList.get(i));
                createQuery.append(Constant.SINGLE_SPACE);
                createQuery.append(fieldTypeList.get(i));
                createQuery.append(Constant.COMMA);
            }
            createQuery.deleteCharAt(createQuery.length() - 1);
            createQuery.append(")");
//            System.out.println(createQuery);

            StringBuilder insertQuery = new StringBuilder();
            insertQuery.append("INSERT INTO ");
            insertQuery.append(sheet);
            insertQuery.append("(");
            for (int i = 1; i < fieldList.size(); i++) {
                insertQuery.append(fieldList.get(i));
                insertQuery.append(Constant.COMMA);
            }
            insertQuery.deleteCharAt(insertQuery.length() - 1);
            insertQuery.append(")");
            insertQuery.append(Constant.SINGLE_SPACE);
            insertQuery.append("VALUES");
            insertQuery.append("(");
            for (int i = 1; i < fieldList.size(); i++) {
                insertQuery.append("?");
                insertQuery.append(Constant.COMMA);
            }
            insertQuery.deleteCharAt(insertQuery.length() - 1);
            insertQuery.append(")");
//            System.out.println(insertQuery);

            Statement stmt;
            PreparedStatement insertPreparedStatement;


            conn.setAutoCommit(false);
            stmt = conn.createStatement();
            stmt.execute(createQuery.toString());
            stmt.close();

            for (int i = 2; i < rowNum; i++) {
                insertPreparedStatement = conn.prepareStatement(insertQuery.toString());
                for (int j = 0; j < colNum; j++) {
                    insertPreparedStatement.setString(j + 1, excelUtil.getCellData(sheet, i, j));
                }
                insertPreparedStatement.executeUpdate();

            }
            conn.commit();
        }
    }
}
