package controller.order;

import model.Orders;

import java.util.List;

public interface OrderService {

String getNextOrderID();

String fetchItemDetailsByItemDescription(String code);

String fetchSupplierDetailsByName(String code);

List<String> fetchItemNamesFromDatabase(String itemName);

boolean placeOrder(Orders orders );


}
