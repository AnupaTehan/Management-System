package single;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class ItemManageForm {

    private static ItemManageForm instance;

    private Stage stage;


    private ItemManageForm(){
        stage = new Stage();
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("../view/itemManageForm.fxml"));
            Parent root = loader.load();
            stage.setScene(new Scene(root));
            stage.setTitle("Item Manager Form");
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public static ItemManageForm getInstance(){
        if (instance== null){
            instance= new ItemManageForm();
        }
        return instance;
    }

    public void show (){
        stage.show();
    }


}
