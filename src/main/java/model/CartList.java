package model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor

public class CartList {

    private String orderId;
    private  String itemId;
    private String itemName;
    private Double unitPrice;
    private Integer quantity;
    private Double total;



}
