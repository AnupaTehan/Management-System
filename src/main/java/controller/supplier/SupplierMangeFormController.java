package controller.supplier;

import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.geometry.Side;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import model.Item;
import model.Supplier;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import single.DashBoardForm;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.ResourceBundle;

public class SupplierMangeFormController implements Initializable {

    @FXML
    private TableColumn colSupplierEmail;

    @FXML
    private TextField txtSupplierEmail;

    @FXML
    private TextField supplierEmail;

    @FXML
    private TextField SupplierAddress;

    @FXML
    private TextField SupplierContactNo;

    @FXML
    private TableColumn  colSupplierAddress;

    @FXML
    private TableColumn  colSupplierContactNo;

    @FXML
    private TableColumn  colSupplierId;

    @FXML
    private TableColumn  colSupplierName;

    @FXML
    private TextField SupplierId;

    @FXML
    private TextField SupplierName;

    @FXML
    private TableView tblSupplier;

    @FXML
    private TextField txtSearchSupplierByName;

    @FXML
    private TextArea txtSupplierAddress;

    @FXML
    private TextField txtSupplierContactNo;

    @FXML
    private TextField txtSupplierId;

    @FXML
    private TextField txtSupplierName;

    SupplierService supplierService = SupplierController.getInstance();

    private ContextMenu suggestionsMenu = new ContextMenu();

    @FXML
    void btnAddSupplierOnAction(ActionEvent event) {

        String contactNo = txtSupplierContactNo.getText();
        boolean state =  true;

        if (!(contactNo.startsWith("07") && contactNo.length()==10)){
            new Alert(Alert.AlertType.ERROR, "Invalid contact Number ! please try again again").show();
            state = false;
            txtSupplierContactNo.setText(null);
        }
        if(state){
            Supplier supplier = new Supplier(
                    txtSupplierId.getText(),
                    txtSupplierName.getText(),
                    txtSupplierAddress.getText(),
                    txtSupplierEmail.getText(),
                    txtSupplierContactNo.getText()
            );
            if(supplierService.addSupplier(supplier)){
                new Alert(Alert.AlertType.INFORMATION , "Supplier Added successfully").show();
                clearFields();
                nextIdGenerated();
            }else{
                new Alert(Alert.AlertType.ERROR,"Supplier not Added ! please try again").show();
            }
        }




    }



    @FXML
    void btnClearOnAction(ActionEvent event) {
        txtSearchSupplierByName.setText(null);
        SupplierId.setText(null);
        SupplierName.setText(null);
        SupplierAddress.setText(null);
        supplierEmail.setText(null);
        SupplierContactNo.setText(null);

    }

    @FXML
    void btnDeleteOnAction(ActionEvent event) {
            if(supplierService.deleteSupplier(SupplierId.getText())){
                new Alert(Alert.AlertType.INFORMATION,"Supplier Deleted !!").show();
                txtSearchSupplierByName.setText(null);
                clearSearchField();
                loadTable();
                nextIdGenerated();
            }else{
                new Alert(Alert.AlertType.ERROR,"Invalid Supplier Id ! please try again ").show();
                clearSearchField();
            }
    }

    @FXML
    void btnRefreshOnAction(ActionEvent event) {
            loadTable();
    }

    @FXML
    void btnSearchOnAction(ActionEvent event) {
        String supplierId = txtSearchSupplierByName.getText();
        Supplier supplier = supplierService.SearchSupplier(supplierId);
        if (Objects.equals(supplier.getSupplierId(), supplierId)){
            SupplierId.setText(supplier.getSupplierId());
            SupplierName.setText(supplier.getSupplierName());
            SupplierAddress.setText(supplier.getSupplierAddress());
            supplierEmail.setText(supplier.getSupplierEmail());
            SupplierContactNo.setText(supplier.getContactNo());
        }else{
            new Alert(Alert.AlertType.ERROR,"Invalid Supplier Id! please try again").show();
            clearSearchField();
        }

    }

    private void clearSearchField() {
        txtSearchSupplierByName.setText(null);
        SupplierId.setText(null);
        SupplierName.setText(null);
        SupplierAddress.setText(null);
        supplierEmail.setText(null);
        SupplierContactNo.setText(null);
    }

    @FXML
    void btnUpdateOnAction(ActionEvent event) {

        Supplier supplier =  new Supplier(
                SupplierId.getText(),
                SupplierName.getText(),
                SupplierAddress.getText(),
                supplierEmail.getText(),
                SupplierContactNo.getText()
        );
        if (supplierService.updateSupplier(supplier)){
            new Alert(Alert.AlertType.INFORMATION, "Supplier Update Successfully").show();
            loadTable();
            clearSearchField();


        }else{
            new Alert(Alert.AlertType.ERROR ,  "Supplier Not Updated! please try again").show();
        }

    }

    @FXML
    void homeMouseOnClick(MouseEvent event) {
        DashBoardForm dashBoardForm = DashBoardForm.getInstance();
        dashBoardForm.show();
    }

    private void clearFields() {
        txtSupplierId.setText(null);
        txtSupplierName.setText(null);
        txtSupplierAddress.setText(null);
        txtSupplierEmail.setText(null);
        txtSupplierContactNo.setText(null);
    }

    public void nextIdGenerated() {
        String lastId = supplierService.getNextSupplierId(); // e.g., "SUP001", "SUP999", etc.

        if (lastId != null && lastId.startsWith("SUP")) {
            try {
                // Extract numeric part after "SUP" (first 3 characters)
                String numericPart = lastId.substring(3);
                int idNum = Integer.parseInt(numericPart);

                // Increment number
                int nextId = idNum + 1;

                // Format with leading zeros (minimum 3 digits, grows automatically)
                String newId = String.format("SUP%0" + Math.max(3, String.valueOf(nextId).length()) + "d", nextId);

                txtSupplierId.setText(newId);

            } catch (NumberFormatException e) {
                // If parsing fails, start fresh
                txtSupplierId.setText("SUP001");
            }
        } else {
            // If no ID is found, start from SUP001
            txtSupplierId.setText("SUP001");
        }
    }



    public void loadTable(){
        ObservableList<Supplier> supplierObservableList = supplierService.getAll();

        if (supplierObservableList.isEmpty()){
            tblSupplier.setItems(supplierObservableList);
        }else{
            tblSupplier.getItems().clear();
            tblSupplier.setItems(supplierObservableList);
        }
    }




    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        colSupplierId.setCellValueFactory(new PropertyValueFactory<>("supplierId"));
        colSupplierName.setCellValueFactory(new PropertyValueFactory<>("supplierName"));
        colSupplierAddress.setCellValueFactory(new PropertyValueFactory<>("supplierAddress"));
        colSupplierEmail.setCellValueFactory(new PropertyValueFactory<>("supplierEmail"));
        colSupplierContactNo.setCellValueFactory(new PropertyValueFactory<>("contactNo"));

        loadTable();

        nextIdGenerated();

        setupArrowKeyNavigation();

        setupSupplierAutoComplete();
    }
    private void setupSupplierAutoComplete() {
        txtSearchSupplierByName.textProperty().addListener((obs, oldText, newText) -> {
            if (newText != null && newText.length() >= 1) showSuggestions(newText);
            else suggestionsMenu.hide();
        });
    }

    private void showSuggestions(String searchText) {
        List<Supplier> suppliers = SupplierController.getInstance().searchSupplierByNamePattern(searchText);
        if (suppliers.isEmpty()) {
            suggestionsMenu.hide();
            return;
        }

        List<MenuItem> menuItems = new ArrayList<>();
        for (Supplier supplier : suppliers) {
            MenuItem item = new MenuItem(supplier.getSupplierName());
            item.setOnAction(e -> {
                SupplierId.setText(supplier.getSupplierId());
                SupplierName.setText(supplier.getSupplierName());
                SupplierAddress.setText(supplier.getSupplierAddress());
                supplierEmail.setText(supplier.getSupplierEmail());
                SupplierContactNo.setText(supplier.getContactNo());
            });
            menuItems.add(item);
        }
        suggestionsMenu.getItems().clear();
        suggestionsMenu.getItems().addAll(menuItems);
        if (!suggestionsMenu.isShowing()) suggestionsMenu.show(txtSearchSupplierByName, Side.BOTTOM, 0, 0);
    }

    private void setupArrowKeyNavigation() {
        // txtSupplierId
        txtSupplierId.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtSupplierName.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtSupplierContactNo.requestFocus(); // wrap-around
                    break;
            }
        });

        // txtSupplierName
        txtSupplierName.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtSupplierAddress.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtSupplierId.requestFocus();
                    break;
            }
        });

        // txtSupplierAddress
        txtSupplierAddress.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtSupplierEmail.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtSupplierName.requestFocus();
                    break;
            }
        });

        // txtSupplierEmail
        txtSupplierEmail.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtSupplierContactNo.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtSupplierAddress.requestFocus();
                    break;
            }
        });

        // txtSupplierContactNo
        txtSupplierContactNo.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtSupplierId.requestFocus(); // wrap-around
                    break;
                case UP:
                case LEFT:
                    txtSupplierEmail.requestFocus();
                    break;
                case ENTER:
                    btnAddSupplierOnAction(new ActionEvent()); // Add supplier on Enter
                    break;
            }
        });
    }


    public void btnExportOnAction(ActionEvent actionEvent) {

        try {
            // Get all items
            ObservableList<Supplier> suppliers = supplierService.getAll();
            if (suppliers.isEmpty()) {
                new Alert(Alert.AlertType.WARNING, "No supplier to export!").show();
                return;
            }

            // Ask user for save location
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Save Excel File");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            fileChooser.setInitialFileName("supplier.xlsx");
            File file = fileChooser.showSaveDialog(tblSupplier.getScene().getWindow());
            if (file == null) return;

            // Create workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Suppliers");

            // === Insert PNG Logo with smaller size ===
            InputStream logoInput = new FileInputStream("src/main/resources/img/logo.png");
            byte[] logoBytes = IOUtils.toByteArray(logoInput);
            logoInput.close();

            int pictureIdx = workbook.addPicture(logoBytes, Workbook.PICTURE_TYPE_PNG);
            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();

            anchor.setCol1(1); // Start column
            anchor.setRow1(1); // Start row
            anchor.setCol2(3); // End column (smaller width)
            anchor.setRow2(8); // End row (smaller height)
            anchor.setDx1(50); // Horizontal offset
            anchor.setDy1(100);  // Vertical offset

            Picture pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize(0.8); // Resize proportionally to fit smaller space

            // === "SUPPLIER LIST" Title with perfect alignment ===
            CellStyle titleStyle = workbook.createCellStyle();
            Font titleFont = workbook.createFont();
            titleFont.setBold(true);
            titleFont.setFontHeightInPoints((short) 16);
            titleFont.setColor(IndexedColors.WHITE.getIndex());
            titleStyle.setFont(titleFont);
            titleStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Add borders for better definition
            titleStyle.setBorderBottom(BorderStyle.MEDIUM);
            titleStyle.setBorderTop(BorderStyle.MEDIUM);
            titleStyle.setBorderLeft(BorderStyle.MEDIUM);
            titleStyle.setBorderRight(BorderStyle.MEDIUM);
            titleStyle.setBottomBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setTopBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setLeftBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setRightBorderColor(IndexedColors.DARK_GREEN.getIndex());

            // Merge cells for title next to logo with proper alignment
            sheet.addMergedRegion(new CellRangeAddress(2, 4, 3, 8)); // Centered vertically with logo

            // Create title row and apply style to all merged cells
            Row titleRow = sheet.getRow(3);
            if (titleRow == null) titleRow = sheet.createRow(2);

            for (int col =3; col <= 3; col++) {
                Cell titleCell = titleRow.createCell(col);
                if (col == 3) {
                    titleCell.setCellValue("SUPPLIER LIST");
                }
                titleCell.setCellStyle(titleStyle);
            }

            // Ensure all rows in the merged region have the style applied
            for (int row = 2; row <= 4; row++) {
                Row currentRow = sheet.getRow(row);
                if (currentRow == null) currentRow = sheet.createRow(row);
                for (int col = 3; col <= 8; col++) {
                    Cell cell = currentRow.getCell(col);
                    if (cell == null) cell = currentRow.createCell(col);
                    cell.setCellStyle(titleStyle);
                }
            }

            // Add spacing rows
            sheet.createRow(5); // Empty row
            sheet.createRow(6); // Empty row

            // === Perfect Table Header Style ===
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Perfect border styling
            headerStyle.setBorderBottom(BorderStyle.MEDIUM);
            headerStyle.setBorderTop(BorderStyle.MEDIUM);
            headerStyle.setBorderLeft(BorderStyle.MEDIUM);
            headerStyle.setBorderRight(BorderStyle.MEDIUM);
            headerStyle.setBottomBorderColor(IndexedColors.DARK_GREEN.getIndex());
            headerStyle.setTopBorderColor(IndexedColors.DARK_GREEN.getIndex());
            headerStyle.setLeftBorderColor(IndexedColors.DARK_GREEN.getIndex());
            headerStyle.setRightBorderColor(IndexedColors.DARK_GREEN.getIndex());

            // Perfect alignment
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

// Add text wrapping for better header display
            headerStyle.setWrapText(true);

            String[] columns = {"Supplier ID", "Supplier Name", "Supplier Address", "Supplier Email", "Contact No"};

// Header row with perfect positioning
            int headerRowIndex = 9;
            int columnOffset = 1; // Start from column B (index 1) instead of A (index 0)
            Row headerRow = sheet.createRow(headerRowIndex);
            headerRow.setHeightInPoints(29); // Set consistent header row height

            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i + columnOffset); // Add offset here
                cell.setCellValue(columns[i]);
                cell.setCellStyle(headerStyle);
            }

// === Table Body Style ===
            CellStyle bodyStyle = workbook.createCellStyle();
            bodyStyle.setBorderBottom(BorderStyle.THIN);
            bodyStyle.setBorderTop(BorderStyle.THIN);
            bodyStyle.setBorderLeft(BorderStyle.THIN);
            bodyStyle.setBorderRight(BorderStyle.THIN);
            bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            bodyStyle.setWrapText(true);

// Fill table data
            int rowNum = headerRowIndex + 1;
            for (Supplier supplier : suppliers) {
                Row row = sheet.createRow(rowNum++);
                row.setHeightInPoints(20); // Consistent row height

                // Create cells with offset
                row.createCell(0 + columnOffset).setCellValue(supplier.getSupplierId());
                row.createCell(1 + columnOffset).setCellValue(supplier.getSupplierName());
                row.createCell(2 + columnOffset).setCellValue(supplier.getSupplierAddress());
                row.createCell(3 + columnOffset).setCellValue(supplier.getSupplierEmail());
                row.createCell(4 + columnOffset).setCellValue(supplier.getContactNo());

                // Apply styles with offset
                for (int i = 0; i < columns.length; i++) {
                    Cell cell = row.getCell(i + columnOffset);
                    if (cell != null) {
                        cell.setCellStyle(bodyStyle);
                    }
                }
            }

// === Set Column Widths with offset ===
            sheet.setColumnWidth(0 + columnOffset, 5000);  // Column B
            sheet.setColumnWidth(1 + columnOffset, 10000); // Column C
            sheet.setColumnWidth(2 + columnOffset, 7000);  // Column D
            sheet.setColumnWidth(3 + columnOffset, 9000);  // Column E
            sheet.setColumnWidth(4 + columnOffset, 6000);  // Column F

            // Save file
            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }
            workbook.close();

            new Alert(Alert.AlertType.INFORMATION, "Excel exported with style!").show();

        } catch (Exception e) {
            e.printStackTrace();
            new Alert(Alert.AlertType.ERROR, "Error: " + e.getMessage()).show();
        }
    }
}

