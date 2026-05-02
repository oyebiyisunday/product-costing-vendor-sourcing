# Product Costing & Vendor Sourcing — Problem Statement

## Context

You are supporting a manufacturing team building a product (for example, a robotic subassembly). Each product is composed of multiple parts sourced from different vendors. Each vendor maintains their own product information and price list.

## Objective

Design an Excel workbook that:

- Defines the parts required to build a product  
- Lets the user select which vendor sources each part  
- Accepts the quantity of each part required  
- Automatically calculates the total cost to build the product  

The solution should be as easy to use and as scalable as practical.

## Workbook structure

### 1. Vendor pricing sheets

Create three worksheets named:

- `Vendor_A`  
- `Vendor_B`  
- `Vendor_C`  

Each vendor sheet must contain a pricing table with at least these columns:

- **Part_ID**  
- **Part_Name**  
- **Unit_Price**  

### 2. Product BOM sheet (final costing)

Create a sheet named **`Product_BOM`**. This is where all final costs must be calculated.

The sheet should support (at minimum) these columns:

| Column            | Description |
|-------------------|-------------|
| Part_ID           | Identifier for the part |
| Part_Name         | Part description |
| Vendor            | One of: Vendor_A, Vendor_B, or Vendor_C |
| Qty_Per_Product   | How many of this part are needed per product |
| Total_Qty         | Calculated quantity |
| Unit_Price        | Pulled dynamically from the selected vendor sheet for that Part_ID |
| Total_Cost        | Calculated cost for the line |

## Requirements

1. **Dynamic pricing** — On `Product_BOM`, unit price for each row must be retrieved based on the selected **Vendor** and **Part_ID**. All pricing must come directly from the vendor sheets (not hard-coded only on the BOM).

2. **Automatic costing** — Costs must update automatically from quantities (including any defined relationship for total quantity vs. per-product quantity).

3. **Deliverable** — Submit the completed Excel workbook with the response (or equivalent delivery).
