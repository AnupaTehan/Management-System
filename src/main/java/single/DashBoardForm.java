package single;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class DashBoardForm {

    private static DashBoardForm instance;

    private Stage stage;


    private DashBoardForm(){
        stage = new Stage();
        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("../view/dashboardForm.fxml"));
            Parent root = loader.load();
            stage.setScene(new Scene(root));
            stage.setTitle("Item Manager Form");
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public static DashBoardForm getInstance(){
        if (instance== null){
            instance= new DashBoardForm();
        }
        return instance;
    }

    public void show (){
        stage.show();
    }
}
