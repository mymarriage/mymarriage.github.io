package ExcelDB;
//E:/JavaWorkSpace/SongsList.xls
//DataReadExcelToSQL

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class DataReadExcelToSQL {
    public static void main(String[] args) {
        Connection connection = null;
        Statement statement = null;
        ResultSet resultSet = null;

        try {
            Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
            connection = DriverManager
                            .getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=E:\\JavaWorkSpace\\SongsList.xlsx;readOnly=false");
            statement = connection.createStatement();
            resultSet = statement.executeQuery("SELECT * FROM [Telugu$]");
            if (resultSet!=null) {
                while (resultSet.next()) {
                	System.out.println(resultSet.getString("SNO"));
                    System.out.println(resultSet.getString("SONGURL"));
                    System.out.println(resultSet.getString("Name"));
                }
            }
        } catch (ClassNotFoundException e) {
            System.err.println("ClassNotFoundException by driver load - " + e);
        } catch (SQLException e) {
            System.err.println("SQLException by connect - " + e);
        } finally {
            if (connection != null) {
                try {
                    connection.close();
                } catch (Exception e) {
                }
            }
            if (statement != null) {
                try {
                    statement.close();
                } catch (Exception e) {
                }
            }
            if (resultSet != null) {
                try {
                    resultSet.close();
                } catch (Exception e) {
                }
            }
        }
    }
}
