package model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor

public class Supplier {

    private String supplierId;
    private String supplierName;
    private String supplierAddress;
    private String supplierEmail;
    private String contactNo;



}
