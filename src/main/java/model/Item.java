package model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor

public class Item {

    private String itemId;
    private String itemName;
    private String unitType;
    private String supplierName;
    private double unitPrice;
    private String date;


}
