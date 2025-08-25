package controller.item;

import controller.supplier.SupplierController;
import controller.supplier.SupplierService;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.geometry.Side;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.MouseEvent;
import model.Item;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import single.DashBoardForm;

import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;

import java.io.*;


public class ItemManageFormController implements Initializable {

    public TableColumn colItemDate;

    public TextField txtItemDate;

    public TextField itemDate;

    public ComboBox cmdItemCodeType;

    public ComboBox cmdSupplierName;


    public ComboBox supplierNameCmd;

    public TableColumn colSupplierName;

    @FXML
    private ComboBox itemUnitType;


    @FXML
    private ComboBox cmdSearchItemType;

    @FXML
    private  TableColumn colItemQuantity;

    @FXML
    private TextField itemQuantity;

    @FXML
    private ComboBox cmdItemType;

    @FXML
    private TableColumn colItemId;

    @FXML
    private TableColumn colItemName;

    @FXML
    private TableColumn colItemType;

    @FXML
    private TableColumn colItemUnitPrice;

    @FXML
    private TextField itemId;

    @FXML
    private TextField itemName;

    @FXML
    private TextField itemType;

    @FXML
    private TextField itemUnitPrice;

    @FXML
    private TableView tblItem;

    @FXML
    private TextField txtItemId;

    @FXML
    private TextField txtItemName;

    @FXML
    private TextField txtItemQuantity;

    @FXML
    private TextField txtItemSearchByName;

    @FXML
    private TextField txtItemUnitPrice;


    ItemService itemservice = ItemController.getInstance();

    SupplierService supplierService = SupplierController.getInstance();

    private ContextMenu itemSuggestionsMenu = new ContextMenu();

    @FXML
    void btnAddItemOnAction(ActionEvent event) {

        Item item = new Item(
                txtItemId.getText(),
                txtItemName.getText(),
                cmdItemType.getValue().toString(),
                cmdSupplierName.getValue().toString(),
                Double.parseDouble(txtItemUnitPrice.getText()),
                txtItemDate.getText()
        );
        if (itemservice.addItem(item)){
            new Alert(Alert.AlertType.INFORMATION, "Item Added Successfully").show();
            clearField();
            setupItemCodeTypeDropdown();

        }else{
            new Alert(Alert.AlertType.ERROR , "Item Not Added ! Please try again").show();
        }



    }

    private void clearField() {
        txtItemId.setText(null);
        txtItemName.setText(null);
        txtItemUnitPrice.setText(null);
        txtItemDate.setText(null);
    }


    @FXML
    void btnClearOnAction(ActionEvent event) {
            clearSearchField();
    }

    private void clearSearchField() {
        itemId.setText(null);
        itemName.setText(null);
        itemUnitPrice.setText(null);
        itemDate.setText(null);

    }

    @FXML
    void btnDeleteOnAction(ActionEvent event) {
        if(itemservice.deleteItem(itemId.getText())){
            new Alert(Alert.AlertType.INFORMATION,"Item Deleted !!").show();
            txtItemSearchByName.setText(null);
            clearSearchField();
            loadTable();

        }else{
            new Alert(Alert.AlertType.ERROR,"Invalid Item  Id ! please try again ").show();
            clearSearchField();
        }

    }


    @FXML
    void btnItemSearchByNameOnAction(ActionEvent event) {
        String searchId = txtItemSearchByName.getText();

        Item item = itemservice.searchItem(searchId);

        if (item != null && Objects.equals(item.getItemId(), searchId)) {
            itemId.setText(item.getItemId());
            itemName.setText(item.getItemName());
            itemUnitType.setValue(item.getUnitType());
            itemUnitPrice.setText(String.valueOf(item.getUnitPrice()));
            itemDate.setText(item.getDate());

        } else {
            new Alert(Alert.AlertType.ERROR, "Invalid Item ID! Please try again.").show();
            clearSearchField();
        }
    }


    @FXML
    void btnRefreshOnAction(ActionEvent event) {
        loadTable();

    }

    @FXML
    void btnUpdateOnAction(ActionEvent event) {
        Item item = new Item(
          itemId.getText(),
          itemName.getText(),
          itemUnitType.getValue().toString(),
                supplierNameCmd.getValue().toString(),
                Double.parseDouble(itemUnitPrice.getText()),
                itemDate.getText()
        );
        if (itemservice.updateItem(item)){
            new Alert(Alert.AlertType.INFORMATION,"Item Updated Successfully").show();
            txtItemSearchByName.setText(null);
            loadTable();
            clearSearchField();
        }else{
            new Alert(Alert.AlertType.ERROR,"Item Not Updated ! please try again ").show();
        }


    }

    @FXML
    void homeMouseOnClick(MouseEvent event) {
        DashBoardForm dashBoardForm = DashBoardForm.getInstance();
        dashBoardForm.show();

    }

    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        colItemId.setCellValueFactory(new PropertyValueFactory<>("itemId"));
        colItemName.setCellValueFactory(new PropertyValueFactory<>("itemName"));
        colItemType.setCellValueFactory(new PropertyValueFactory<>("unitType"));
        colSupplierName.setCellValueFactory(new PropertyValueFactory<>("supplierName"));
        colItemUnitPrice.setCellValueFactory(new PropertyValueFactory<>("unitPrice"));
        colItemDate.setCellValueFactory(new PropertyValueFactory<>("date"));

        ObservableList<String> itemTypeList = FXCollections.observableArrayList("Each", "Pair");
        cmdItemType.setItems(itemTypeList);
        itemUnitType.setItems(itemTypeList);

        ObservableList<String> ItemCodeTypeList = FXCollections.observableArrayList("Head Protection", "Eye Protection" , "Ear Protection","Respirator and Mask","Hand Protection","Body Protection","Foot Protection");
        cmdItemCodeType.setItems(ItemCodeTypeList);
        setupItemCodeTypeDropdown();


        loadTable();

        setupArrowKeyNavigationForItemForm();
        setupItemAutoComplete();


        txtItemDate.setText(LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));



        cmdSupplierName.setEditable(true);

// Load all suppliers at the beginning
        ObservableList<String> allSuppliers = FXCollections.observableArrayList(supplierService.getSuppliersByName(""));
        cmdSupplierName.setItems(allSuppliers);

// Add listener for typing
        cmdSupplierName.getEditor().textProperty().addListener((obs, oldValue, newValue) -> {
            if (newValue == null || newValue.trim().isEmpty()) {
                // Show all suppliers if nothing typed
                cmdSupplierName.setItems(allSuppliers);
            } else {
                // Filter the list based on typed text
                List<String> filtered = supplierService.getSuppliersByName(newValue);
                cmdSupplierName.setItems(FXCollections.observableArrayList(filtered));
            }

            // Keep the text in editor and show dropdown
            cmdSupplierName.getEditor().setText(newValue); // editor text
            cmdSupplierName.getEditor().positionCaret(newValue.length()); // move cursor to end
            cmdSupplierName.show();
        });

        supplierNameCmd.setEditable(true);

// Load all suppliers at the beginning
        ObservableList<String> allSuppliersName = FXCollections.observableArrayList(supplierService.getSuppliersByName(""));
        supplierNameCmd.setItems(allSuppliersName);

// Add listener for typing
        supplierNameCmd.getEditor().textProperty().addListener((obs, oldValue, newValue) -> {
            if (newValue == null || newValue.trim().isEmpty()) {
                // Show all suppliers if nothing typed
                supplierNameCmd.setItems(allSuppliersName);
            } else {
                // Filter the list based on typed text
                List<String> filtered = supplierService.getSuppliersByName(newValue);
                supplierNameCmd.setItems(FXCollections.observableArrayList(filtered));
            }

            // Keep the text in editor and show dropdown
            supplierNameCmd.getEditor().setText(newValue); // editor text
            supplierNameCmd.getEditor().positionCaret(newValue.length()); // move cursor to end
            supplierNameCmd.show();
        });



    }
    private void setupItemAutoComplete() {
        txtItemSearchByName.textProperty().addListener((obs, oldText, newText) -> {
            if (newText != null && newText.length() >= 1) showItemSuggestions(newText);
            else itemSuggestionsMenu.hide();
        });
    }

    private void showItemSuggestions(String searchText) {
        List<Item> items = ItemController.getInstance().searchItemsByNamePattern(searchText);
        if (items.isEmpty()) {
            itemSuggestionsMenu.hide();
            return;
        }

        List<MenuItem> menuItems = new ArrayList<>();
        for (Item item : items) {
            MenuItem menuItem = new MenuItem(item.getItemName());
            menuItem.setOnAction(e -> {
                itemId.setText(item.getItemId());
                itemName.setText(item.getItemName());
                itemUnitType.setValue(item.getUnitType());
                supplierNameCmd.setValue(item.getSupplierName());
                itemUnitPrice.setText(String.valueOf(item.getUnitPrice()));
                itemDate.setText(item.getDate());

                itemSuggestionsMenu.hide();
            });
            menuItems.add(menuItem);
        }
        itemSuggestionsMenu.getItems().clear();
        itemSuggestionsMenu.getItems().addAll(menuItems);
        if (!itemSuggestionsMenu.isShowing()) itemSuggestionsMenu.show(txtItemSearchByName, Side.BOTTOM, 0, 0);
    }

    private void loadTable() {
        ObservableList<Item> itemObservableList = itemservice.getAll();

        if (itemObservableList.isEmpty()){
            tblItem.setItems(itemObservableList);
        }else{
            tblItem.getItems().clear();
            tblItem.setItems(itemObservableList);
        }
    }







    private void setupArrowKeyNavigationForItemForm() {
        // txtItemId
        txtItemId.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtItemName.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtItemDate.requestFocus(); // wrap-around
                    break;
            }
        });

        // txtItemName
        txtItemName.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    cmdItemType.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtItemId.requestFocus();
                    break;
            }
        });

        // cmdItemType
        cmdItemType.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtItemUnitPrice.requestFocus();
                    break;
                case UP:
                case LEFT:
                    txtItemName.requestFocus();
                    break;
            }
        });

        // txtItemUnitPrice
        txtItemUnitPrice.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtItemDate.requestFocus();
                    break;
                case UP:
                case LEFT:
                    cmdItemType.requestFocus();
                    break;
            }
        });

        // txtItemDate
        txtItemDate.setOnKeyPressed(event -> {
            switch (event.getCode()) {
                case DOWN:
                case RIGHT:
                    txtItemId.requestFocus(); // wrap-around
                    break;
                case UP:
                case LEFT:
                    txtItemUnitPrice.requestFocus();
                    break;
                case ENTER:
                    btnAddItemOnAction(new ActionEvent()); // Add item on Enter
                    break;
            }
        });
    }



    private void setupItemCodeTypeDropdown() {
        ObservableList<String> itemCodeTypeList = FXCollections.observableArrayList(
                "Head Protection",
                "Eye Protection",
                "Ear Protection",
                "Respirator and Mask",
                "Hand Protection",
                "Body Protection",
                "Foot Protection",
                "Others"
        );
        cmdItemCodeType.setItems(itemCodeTypeList);

        // Map category name to prefix
        Map<String, String> categoryPrefixes = Map.of(
                "Head Protection", "HEP",
                "Eye Protection", "EYP",
                "Ear Protection", "EAR",
                "Respirator and Mask", "RAM",
                "Hand Protection", "HAP",
                "Body Protection", "BOP",
                "Foot Protection", "FOP",
                "Others", "OTH"
        );

        // Generate ID when a category is selected
        cmdItemCodeType.setOnAction(e -> {
            String selectedCategory = (String) cmdItemCodeType.getValue();
            if (selectedCategory != null) {
                String prefix = categoryPrefixes.get(selectedCategory);
                String nextId = itemservice.getNextItemId(prefix);
                txtItemId.setText(nextId);
            }
        });
    }




    public void btnExportOnAction(ActionEvent actionEvent) {
        try {
            // Get all items
            ObservableList<Item> items = itemservice.getAll();
            if (items.isEmpty()) {
                new Alert(Alert.AlertType.WARNING, "No items to export!").show();
                return;
            }

            // Ask user for save location
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Save Excel File");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            fileChooser.setInitialFileName("Items.xlsx");
            File file = fileChooser.showSaveDialog(tblItem.getScene().getWindow());
            if (file == null) return;

            // Create workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Items");

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

            // === "ITEM LIST" Title with perfect alignment ===
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

            for (int col = 3; col <= 3; col++) {
                Cell titleCell = titleRow.createCell(col);
                if (col == 3) {
                    titleCell.setCellValue("ITEM LIST");
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

            String[] columns = {"Item ID", "Item Name", "Unit Type", "Unit Price", "Date"};

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
            for (Item item : items) {
                Row row = sheet.createRow(rowNum++);
                row.setHeightInPoints(20); // Consistent row height

                // Create cells with offset
                row.createCell(0 + columnOffset).setCellValue(item.getItemId());
                row.createCell(1 + columnOffset).setCellValue(item.getItemName());
                row.createCell(2 + columnOffset).setCellValue(item.getUnitType());
                row.createCell(3 + columnOffset).setCellValue(item.getUnitPrice());
                row.createCell(4 + columnOffset).setCellValue(item.getDate());

                // Apply styles with offset
                for (int i = 0; i < columns.length; i++) {
                    Cell cell = row.getCell(i + columnOffset);
                    if (cell != null) {
                        cell.setCellStyle(bodyStyle);
                    }
                }
            }

            // === Set Column Widths with offset ===
            sheet.setColumnWidth(0 + columnOffset, 5000);  // Column B - Item ID
            sheet.setColumnWidth(1 + columnOffset, 10000); // Column C - Item Name
            sheet.setColumnWidth(2 + columnOffset, 7000);  // Column D - Unit Type
            sheet.setColumnWidth(3 + columnOffset, 7000);  // Column E - Unit Price
            sheet.setColumnWidth(4 + columnOffset, 8000);  // Column F - Date

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


