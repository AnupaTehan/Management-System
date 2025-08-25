package single;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class OrderManageForm {

    private static OrderManageForm instance;

    private Stage stage;


    private OrderManageForm(){
        stage = new Stage();
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("../view/OrderManagementForm.fxml"));
            Parent root = loader.load();
            stage.setScene(new Scene(root));
            stage.setTitle("Order Manager Form");
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public static OrderManageForm getInstance(){
        if (instance== null){
            instance= new OrderManageForm();
        }
        return instance;
    }

    public void show (){
        stage.show();
    }
}
