# 📦 Excel-SAP Delivery Date Updater (VBA Macro)

This project contains a VBA macro for Microsoft Excel that automates the update of delivery dates in **SAP** using **SAP GUI scripting**. It's especially useful for logistics, supply chain, or operations teams working with high volumes of deliveries that need regular rescheduling or alignment.

---

## 🧠 Features

- Reads delivery numbers and new dates from an Excel sheet.
- Automatically logs into SAP and opens transaction **VL02N**.
- Updates key date fields including:
  - Picking Date
  - Transportation Planning Date
  - Loading Date
  - Planned Goods Issue Date
  - Delivery Date
- Logs "Updated" in Excel for each successful record.

---

## 📁 File Structure

- `DeliveryUpdater.xlsm`: Excel workbook with the embedded macro.
- `README.md`: Project description and usage guide.
- `UpdateDeliveryDates.bas` *(optional)*: Exported macro file (for easy import if needed).

---

## 📋 Requirements

- **Microsoft Excel** (with macros enabled)
- **SAP GUI** with scripting enabled
- Valid SAP login credentials and access to transaction **VL02N**
- Excel file named `Deliveries` with:
  - **Column A:** Delivery Number
  - **Column B:** New Delivery Date
  - **Column C:** (Used by macro to indicate "Updated")

---

## 🛠️ How to Use

1. **Open Excel** and enable macros.
2. Go to the **"Deliveries"** sheet.
3. Fill in:
   - Column A: SAP Delivery Numbers (e.g., 80001234)
   - Column B: Desired new delivery dates (any date format; macro will reformat)
4. Press `Alt + F8`, select `UpdateDeliveryDates`, and click **Run**.
5. Wait for the macro to loop through all entries. A message box will confirm completion.

> 📝 Each updated line will be marked as `"Updated"` in Column C.

---

## ⚠️ Notes

- This script interacts directly with the SAP GUI. Make sure:
  - No SAP pop-ups interrupt the process.
  - You do not interact with the keyboard or mouse during execution.
- Always test with a few deliveries before running it on large volumes.

---

## ✅ Example Use Case

You received a customer request to delay 200 shipments by one week. Instead of updating them one by one in SAP, paste them into the Excel sheet, run the macro, and it will handle all the changes automatically.

---

## 💬 Contact

For questions or suggestions, feel free to connect via: lauralizethprado22@gmail.com

