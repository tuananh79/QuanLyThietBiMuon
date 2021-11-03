package KetNoiSQL;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


import java.sql.Connection;
import java.sql.DriverManager;
import  java.sql.*;

/**
 *
 * @author Minh Huan
 */
public class KetNoi {

    /**
     * @param args the command line arguments
     * @return 
     */
    public static Connection ConnectSQL(){
        Connection ketNoi = null;
        String url = "jdbc:sqlserver://;databaseName=QLTB";
        String userName = "sa";
        String password = "123";
        
        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            ketNoi = (Connection) DriverManager.getConnection(url, userName, password);
            System.out.println("Connect sucess...");
        } catch (Exception ex) {
            System.out.println("Connect error...");
        }
        return ketNoi;
    }
    
    
    public static void main(String[] args) {
        // TODO code application logic here
    }
    
}
