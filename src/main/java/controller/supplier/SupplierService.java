package controller.supplier;

import javafx.collections.ObservableList;
import model.Supplier;

import java.util.List;

public interface SupplierService {

    boolean addSupplier(Supplier supplier);

    boolean updateSupplier(Supplier supplier);

    boolean deleteSupplier(String supplierId);

    Supplier SearchSupplier(String supplierId);

    ObservableList<Supplier> getAll();

    String getNextSupplierId();

    ObservableList<String> getSupplierIds();


    List<Supplier> searchSupplierByNamePattern(String supplierNamePart);

    List<String> getSuppliersByName(String filter);
}
