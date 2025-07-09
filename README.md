# Ecommerce-Inventory-VBA-System

## Overview

**Ecommerce-Inventory-VBA-System** is a complete Excel-based Inventory and Sales Management System built using **VBA (Visual Basic for Applications)**. This project is ideal for small businesses, e-commerce sellers, and inventory managers looking for a lightweight solution without the need for external tools or add-ins.

It features a clean UI, color-coded inventory indicators, real-time stock tracking, and a modular code structure designed for easy maintenance and scalability.


## Features

### Inventory Tracking
- Real-time stock level monitoring
- Automatic update of stock after each sale
- Color-coded alerts for stock levels:
  - **Red**: Low stock
  - **Gray**: Out of stock
  - **Green**: Sufficient stock

### Product Management
- Add, edit, and delete product details
- Track categories, prices, stock levels, and reorder points
- UserForm for easy product entry and management

### Sales Module
- Process sales and automatically deduct sold quantities
- ComboBox to select products quickly
- Input validation and error handling

### Dashboard Analytics
- Overview of total products
- List of low-stock and out-of-stock items
- Daily sales totals and trends

### Visual UI
- Light blue-gray background for forms
- Readable labels and layout
- Color-coded rows in the inventory sheet based on stock status


## How to Use

### Step 1: Download & Enable Macros
- Download the file and open it in Excel
- Save it as `.xlsm` (macro-enabled format)
- Enable macros when prompted

### Step 2: Initialize the System
1. Press `Alt + F11` to open the VBA Editor
2. Go to `Insert > Module` and paste the provided code
3. Press `Ctrl + G` to open the Immediate Window
4. Run the following:
   ```vba
   QuickSystemSetup
## Step 3: Create UserForms

### Product Management Form (`frmProductManagement`)

**Controls:**
- **TextBoxes:**
  - `txtProductID`
  - `txtProductName`
  - `txtCategory`
  - `txtInitialStock`
  - `txtReorderLevel`
  - `txtUnitPrice`
- **Buttons:**
  - `cmdAddProduct`
  - `cmdClose`

### Sales Processing Form (`frmSalesProcessing`)

**Controls:**
- **ComboBox:**
  - `cmbProductID`
- **TextBox:**
  - `txtQuantity`
- **Buttons:**
  - `cmdProcessSale`
  - `cmdClose`

## Commands

Use these commands in the VBA Immediate Window (`Ctrl + G`):

- `QuickSystemSetup` – Initializes the inventory structure
- `RefreshDashboard` – Updates the summary stats
- `ColorCodeInventory` – Applies color coding to inventory
- `GenerateInventoryReport` – Creates a printable report

## Code Highlights

- Modular structure (separate functions for setup, UI, processing)
- Commented code for better readability
- Descriptive naming conventions
- Full error handling
- No external dependencies or add-ins

## Requirements

- Microsoft Excel 2016 or later
- Macros enabled
- Basic VBA knowledge (for customization)

## Contact

**Developer**: Qamar Mehmood  
**Email**: qamarmehmood533@gmail.com  
**Phone/WhatsApp**: +92 349 7949757
