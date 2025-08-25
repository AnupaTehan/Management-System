package single;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class SupplierManageForm {

    private static SupplierManageForm  instance;

    private Stage stage;


    private SupplierManageForm(){
        stage = new Stage();
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("../view/supplierManageForm.fxml"));
            Parent root = loader.load();
            stage.setScene(new Scene(root));
            stage.setTitle("Supplier Manager Form");
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public static SupplierManageForm getInstance(){
        if (instance== null){
            instance= new SupplierManageForm();
        }
        return instance;
    }

    public void show (){
        stage.show();
    }
}
