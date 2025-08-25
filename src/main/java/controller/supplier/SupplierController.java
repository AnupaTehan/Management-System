package controller.supplier;

import db.DBConnection;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import model.Supplier;
import util.CrudUtil;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class SupplierController implements SupplierService {


    private static SupplierController instance;

    private SupplierController(){
    }

    public static SupplierController getInstance(){
        return instance == null  ? instance = new SupplierController() : instance;
    }


    @Override
    public boolean addSupplier(Supplier supplier) {
        String SQL  = "INSERT INTO supplier VALUES(?,?,?,?,?)";

        try{
            return CrudUtil.execute(
                    SQL,
                    supplier.getSupplierId(),
                    supplier.getSupplierName(),
                    supplier.getSupplierAddress(),
                    supplier.getSupplierEmail(),
                    supplier.getContactNo()
            );
        }catch (SQLException e){
            throw new RuntimeException(e);
        }
    }

    @Override
    public boolean updateSupplier(Supplier supplier) {
        String SQL = "UPDATE supplier SET supplierName = ?, supplierAddress = ?, contactNo = ?, supplierEmail = ? WHERE supplierId = ?";
        try {
            return CrudUtil.execute(
                    SQL,
                    supplier.getSupplierName(),      // 1st placeholder
                    supplier.getSupplierAddress(),   // 2nd
                    supplier.getContactNo(),         // 3rd
                    supplier.getSupplierEmail(),     // 4th
                    supplier.getSupplierId()         // 5th (WHERE)
            );
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }



    @Override
    public boolean deleteSupplier(String supplierId) {
        String SQL = "DELETE FROM supplier WHERE supplierId = ?";
        try {
            return CrudUtil.execute(SQL, supplierId);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }


    @Override
    public Supplier SearchSupplier(String supplierId) {
       String SQL = "SELECT * FROM  supplier WHERE supplierId = ?";

       try {
           ResultSet resultSet = CrudUtil.execute(SQL , supplierId);
           while(resultSet.next()){
               return new Supplier(
                       resultSet.getString(1),
                       resultSet.getString(2),
                       resultSet.getString(3),
                       resultSet.getString(4),
                       resultSet.getString(5)
               );
           }
       } catch (SQLException e) {
           throw new RuntimeException(e);
       }
       return null;
    }

    @Override
    public ObservableList<Supplier> getAll() {
        ObservableList<Supplier> supplierObservableList = FXCollections.observableArrayList();
        String SQL = "SELECT * FROM supplier";
        try{
            ResultSet resultSet = CrudUtil.execute(SQL);
            while(resultSet.next()){
                supplierObservableList.add(new Supplier(
                        resultSet.getString(1),
                        resultSet.getString(2),
                        resultSet.getString(3),
                        resultSet.getString(4),
                        resultSet.getString(5)
                ));

            }
            return supplierObservableList;


        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public String getNextSupplierId() {
        String SQL = "SELECT supplierId FROM supplier ORDER BY supplierId DESC LIMIT 1 ";
        try{
            Connection connection = DBConnection.getInstance().getConnection();
            PreparedStatement pstm = connection.prepareStatement(SQL);
            ResultSet resultSet = pstm.executeQuery();
            String lastID ="";

            if (resultSet.next()){
                lastID = resultSet.getString("supplierId");
            }else{
                lastID = "SUP000";
            }
            return lastID;

        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public ObservableList<String> getSupplierIds() {
        ObservableList <String> suppliersId =  FXCollections.observableArrayList();
        ObservableList<Supplier>supplierObservableList = getAll();
        supplierObservableList.forEach(supplier -> {
            suppliersId.add(supplier.getSupplierId());
        });


        return suppliersId;
    }

    @Override
    public List<Supplier> searchSupplierByNamePattern(String supplierNamePart) {
        String SQL = "SELECT * FROM supplier WHERE supplierName LIKE ?";

        List<Supplier> suppliers = new ArrayList<>();

        try {
            Connection connection = DBConnection.getInstance().getConnection();
            PreparedStatement pstm = connection.prepareStatement(SQL);
            pstm.setString(1, supplierNamePart + "%"); // starts with
            ResultSet rs = pstm.executeQuery();

            while (rs.next()) {
                suppliers.add(new Supplier(
                        rs.getString("supplierId"),
                        rs.getString("supplierName"),
                        rs.getString("supplierAddress"),
                        rs.getString("supplierEmail"),
                        rs.getString("contactNo")
                ));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return suppliers;
    }

    @Override
    public List<String> getSuppliersByName(String filter) {
        List<String> supplierNames = new ArrayList<>();
        String sql;

        if (filter == null || filter.trim().isEmpty()) {
            sql = "SELECT supplierName FROM supplier"; // show all
        } else {
            sql = "SELECT supplierName FROM supplier WHERE supplierName LIKE ?";
        }

        // Use a NEW connection for each query
        try (Connection connection = DBConnection.getNewConnection();
             PreparedStatement ps = connection.prepareStatement(sql)) {

            if (filter != null && !filter.trim().isEmpty()) {
                ps.setString(1, "%" + filter + "%");
            }

            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    supplierNames.add(rs.getString("supplierName"));
                }
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return supplierNames;
    }




}
