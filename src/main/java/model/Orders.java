package model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor


public class Orders {
 private String orderId;
 private String orderDate;
 private String supplierId;
 private String supplierName;
 private String contactNo;
 private String supplierAddress;
 private String supplierEmail;
 private Double netTotal;
List <CartList> cartLists;

}
