package com.sovos;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class Conexao {
    Connection con = null;


    public Conexao() {
    }

    public Connection getNewConnection()
    {

        try {
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/desafiosql", "root", "99141473");

        }   catch(SQLException ex) {
            System.out.println("SQLException:" + ex.getMessage());

        }    catch (Exception e)
        {
            System.out.println(e.getMessage());

        }
        return con;
    }

    public Connection getNewConecction() {
        try {
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/desafiosql", "root", "99141473");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return con;
    }
}
