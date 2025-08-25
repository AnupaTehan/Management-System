package controller;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import single.ItemManageForm;
import single.OrderManageForm;
import single.SupplierManageForm;

public class DashboardFormController {

    @FXML
    void btnItemManagementOnAction(ActionEvent event) {
        ItemManageForm itemManageForm = ItemManageForm.getInstance();
        itemManageForm.show();

    }

    @FXML
    void btnPurchaseOrderOnAction(ActionEvent event) {
       OrderManageForm orderManageForm= OrderManageForm.getInstance();
        orderManageForm.show();

    }

    @FXML
    void btnSupplierManagementOnAction(ActionEvent event) {
        SupplierManageForm supplierManageForm = SupplierManageForm.getInstance();
        supplierManageForm.show();
    }



}
