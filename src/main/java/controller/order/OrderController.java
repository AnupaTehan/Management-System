package controller.order;

import controller.supplier.SupplierController;
import db.DBConnection;
import model.Orders;
import util.CrudUtil;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

public class OrderController implements OrderService {


    private static OrderController instance;

    private OrderController(){
    }

    public static OrderController getInstance(){
        return instance == null  ? instance = new OrderController() : instance;
    }

    @Override
    public String getNextOrderID() {
        String SQL = "SELECT orderId FROM orders ORDER BY orderId DESC LIMIT 1";
        try{
            Connection connection = DBConnection.getInstance().getConnection();
            PreparedStatement pstm = connection.prepareStatement(SQL);
            ResultSet resultSet = pstm.executeQuery();
            String lastID ="";

            if (resultSet.next()){
                lastID = resultSet.getString("orderId");
            }else{
                lastID = "PO00";
            }
            return lastID;

        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public String fetchItemDetailsByItemDescription(String code) {
        return "";
    }

    @Override
    public String fetchSupplierDetailsByName(String code) {
        return "";
    }

    @Override
    public List<String> fetchItemNamesFromDatabase(String itemName) {
        return List.of();
    }

    @Override
    public boolean placeOrder(Orders orders) {
        String SQL = "INSERT INTO orders VALUE(?,?,?,?,?,?,?,?)";
        try{
            return CrudUtil.execute(
                    SQL,
                    orders.getOrderId(),
                    orders.getOrderDate(),
                    orders.getSupplierId(),
                    orders.getSupplierName(),
                    orders.getContactNo(),
                    orders.getSupplierAddress(),
                    orders.getSupplierEmail(),
                    orders.getNetTotal()
            );
        }catch (SQLException e){
            throw new RuntimeException(e);
        }
    }
}
