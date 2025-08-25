package db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DBConnection {

    private static DBConnection instance;
    private Connection connection;

    // Define DB properties
    private static final String URL = "jdbc:mysql://localhost:3306/mycustomdb";
    private static final String USER = "root";
    private static final String PASSWORD = "12345";

    // Private constructor (singleton)
    private DBConnection() throws SQLException {
        connection = DriverManager.getConnection(URL, USER, PASSWORD);
    }

    // Return the same connection (singleton)
    public Connection getConnection() {
        return connection;
    }

    // If you want NEW connection every time
    public static Connection getNewConnection() throws SQLException {
        return DriverManager.getConnection(URL, USER, PASSWORD);
    }

    // Get singleton instance
    public static DBConnection getInstance() throws SQLException {
        if (instance == null || instance.getConnection().isClosed()) {
            instance = new DBConnection();
        }
        return instance;
    }
}
