package controller.item;

import db.DBConnection;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import model.Item;
import util.CrudUtil;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class ItemController implements ItemService {

    private static ItemController instance;

    private ItemController(){
    }

    public static ItemController getInstance(){
        return instance == null  ? instance = new ItemController() : instance;
    }


    @Override
    public boolean addItem(Item item) {

        String SQL = "INSERT INTO item VALUES(?,?,?,?,?,?)";
        try{
            return CrudUtil.execute(
                    SQL,
                    item.getItemId(),
                    item.getItemName(),
                    item.getUnitType(),
                    item.getSupplierName(),
                    item.getUnitPrice(),
                    item.getDate()
            );


        }catch (SQLException e){
            throw new RuntimeException(e);
        }
    }

    @Override
    public boolean updateItem(Item item) {
        String SQL = "UPDATE item SET itemName = ?, unitType = ?, supplierName = ? , unitPrice = ?, date = ? WHERE itemId = ?";
        try {
            return CrudUtil.execute(
                    SQL,
                    item.getItemName(),
                    item.getUnitType(),
                    item.getSupplierName(),
                    item.getUnitPrice(),
                    item.getDate(), // Or use java.sql.Date.valueOf(item.getDate())
                    item.getItemId()
            );
        } catch (SQLException e) {
            e.printStackTrace();
            return false;
        }
    }


    @Override
    public boolean deleteItem(String itemId) {
        String SQL = "DELETE FROM item WHERE itemId = ?";
        try {
            return CrudUtil.execute(SQL, itemId);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }


    @Override
    public Item searchItem(String itemId) {

        String SQL = "SELECT * FROM item WHERE itemId=?";
        try{
            ResultSet resultSet = CrudUtil.execute(SQL , itemId);
            while(resultSet.next()){
                return new Item(
                        resultSet.getString(1),
                        resultSet.getString(2),
                        resultSet.getString(3),
                        resultSet.getString(4),
                        resultSet.getDouble(5),
                        resultSet.getString(6)

                );
            }

        }catch (SQLException e){
            throw  new RuntimeException(e);
        }
        return null;

    }

    @Override
    public ObservableList<Item> getAll() {

        ObservableList<Item>itemObservableList = FXCollections.observableArrayList();
       String SQL= "SELECT * FROM item";
       try{
           ResultSet resultSet = CrudUtil.execute(SQL);
           while(resultSet.next()){
               itemObservableList.add(new Item(
                       resultSet.getString(1),
                       resultSet.getString(2),
                       resultSet.getString(3),
                       resultSet.getString(4),
                       resultSet.getDouble(5),
                       resultSet.getString(6)
               ));
           }
           return itemObservableList;
       }catch (SQLException e){
           throw  new RuntimeException(e);
       }
    }

    @Override
    public String getNextItemId(String prefix) {
        String SQL = "SELECT itemId FROM item WHERE itemId LIKE ? ORDER BY itemId DESC LIMIT 1";
        try {
            Connection connection = DBConnection.getInstance().getConnection();
            PreparedStatement pstm = connection.prepareStatement(SQL);
            pstm.setString(1, prefix + "%");

            ResultSet rs = pstm.executeQuery();
            if (rs.next()) {
                String lastId = rs.getString("itemId");
                int lastNum = Integer.parseInt(lastId.substring(prefix.length()));
                int nextNum = lastNum + 1;

                // Dynamically adjust zero padding (minimum 4 digits, grows automatically)
                String newId = prefix + String.format("%0" + Math.max(4, String.valueOf(nextNum).length()) + "d", nextNum);
                return newId;

            } else {
                return prefix + "0001";
            }
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }




    @Override
    public ObservableList<String> getItemIds() {
        ObservableList<String> itemCodes = FXCollections.observableArrayList();
        ObservableList<Item> itemObservableList = getAll();
        itemObservableList.forEach(item -> {
            itemCodes.add(item.getItemId());
        });
        return itemCodes;
    }

    @Override
    public List<Item> searchItemsByNamePattern(String itemNamePart) {
        String SQL = "SELECT * FROM item WHERE itemName LIKE ?";

        List<Item> items = new ArrayList<>();

        try {
            Connection connection = DBConnection.getInstance().getConnection();
            PreparedStatement pstm = connection.prepareStatement(SQL);
            pstm.setString(1, itemNamePart + "%"); // Matches items starting with itemNamePart
            ResultSet rs = pstm.executeQuery();

            while (rs.next()) {
                items.add(new Item(
                        rs.getString("itemId"),
                        rs.getString("itemName"),
                        rs.getString("unitType"),
                        rs.getString("supplierName"),
                        rs.getDouble("unitPrice"),
                        rs.getString("date")
                ));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return items;
    }

    @Override
    public List<String> getItemsBySupplierName(String supplierName) {
        List<String> items = new ArrayList<>();
        String sql = "SELECT itemName FROM Item WHERE supplierName = ?";

        try (Connection connection = DBConnection.getInstance().getConnection();
             PreparedStatement ps = connection.prepareStatement(sql)) {

            ps.setString(1, supplierName);  // Set the supplier name parameter
            ResultSet rs = ps.executeQuery();

            while (rs.next()) {
                items.add(rs.getString("itemName")); // Add each item name to the list
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return items; // Return the list of item names
    }

    @Override
    public List<String> searchItemsNamesByNamePattern(String filter) {
        List<String> itemNames = new ArrayList<>();
        String sql;

        if (filter == null || filter.trim().isEmpty()) {
            sql = "SELECT itemName FROM item"; // show all items
        } else {
            sql = "SELECT itemName FROM item WHERE itemName LIKE ?";
        }

        try (Connection connection = DBConnection.getNewConnection();
             PreparedStatement ps = connection.prepareStatement(sql)) {

            if (filter != null && !filter.trim().isEmpty()) {
                ps.setString(1, "%" + filter + "%");
            }

            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    itemNames.add(rs.getString("itemName"));
                }
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return itemNames;
    }



}
