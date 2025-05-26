# Export To Excel PCF Control

A fully configurable Power Apps Component Framework (PCF) control for exporting JSON data from a Canvas app to an Excel (.xlsx) file. This is the most feature-rich open-source export-to-Excel PCF control, offering extensive styling, icon customization, and column auto-sizing options.

## Features

* **Data Export**: Binds to any JSON string (via `DataToExport`) and converts it into an Excel worksheet.
* **Customizable File & Sheet Names**: Configure `ExportFileName` and `ExportSheetName` properties.
* **Icon Mode**: Toggle between a full button or an icon-only button using the `ShowIcon` property.
* **Multiple Icon Choices**: Choose between a Download icon or an Excel Document icon with `IconName`.
* **Icon Styling**: Adjust `IconColor`, `IconSize`, hover/active icon colors.
* **Auto-Width Columns**: Enable `AutoWidthColumns` to automatically size columns to fit content.
* **Typography & Layout**:

  * Button text font family (`ButtonFont`)
  * Font size (`ButtonTextSize`)
  * Font weight (`FontWeight`) with options from Thin (100) to Black (900)
  * Font style (`FontStyle`): Normal, Italic, Oblique
  * Text decoration (`TextDecoration`): None, Underline, Overline, Line-through
* **Button Dimensions & Spacing**:

  * Width (`ButtonWidth`), Height (`ButtonHeight`)
  * Padding (`Padding`), Margin (`Margin`)
* **Borders & Corners**:

  * Border color (`BorderColor`), width (`BorderWidth`), style (`BorderStyle`)
  * Corner radius (`ButtonRadius`)
* **Color Customization**:

  * Normal state: `ButtonBackgroundColor`, `ButtonTextColor`, `BorderColor`
  * Hover state: `HoverBackgroundColor`, `HoverTextColor`, `HoverBorderColor`
  * Active state: `ActiveBackgroundColor`, `ActiveTextColor`
  * Focus state: `FocusBorderColor`
  * Disabled state: `DisabledBackgroundColor`, `DisabledTextColor`
* **Tooltip Support**: Set `ToolTip` text to show on hover.
* **Office Fabric Integration**: Uses Fabric icons and styles for a native Power Apps look.

## Installation

1. Clone or download this repository.
2. In your PCF project, place `ControlManifest.input.xml` and `index.ts` in the `src` folder.
3. Run the standard PCF build commands:

   ```bash
   npm install
   npm run build
   ```
4. Import the control into your Canvas app via the Power Apps component library.

## Properties Reference

| Property                | Type            | Default                         | Description                                                   |
| ----------------------- | --------------- | ------------------------------- | ------------------------------------------------------------- |
| DataToExport            | SingleLine.Text | **(Required)**                  | JSON string array to export.                                  |
| ExportFileName          | SingleLine.Text | `Exported_Excel_File`           | Filename (without extension) for the downloaded `.xlsx` file. |
| ExportSheetName         | SingleLine.Text | `Exported_Data`                 | Name of the Excel worksheet.                                  |
| ButtonText              | SingleLine.Text | `Export to Excel`               | Text displayed on the button.                                 |
| ShowIcon                | TwoOptions      | `false`                         | Toggle icon-only mode.                                        |
| IconName                | Enum            | `Download`                      | Choose `Download` or `ExcelDocument` icon.                    |
| IconColor               | SingleLine.Text | `RGBA(0, 120, 212, 1)`          | Color of the icon in normal state.                            |
| IconSize                | SingleLine.Text | `14px`                          | Font size for the icon.                                       |
| AutoWidthColumns        | TwoOptions      | `false`                         | Auto-adjust column widths to fit cell content.                |
| ButtonFont              | SingleLine.Text | `Segoe UI`                      | Font family for button text.                                  |
| ButtonTextSize          | SingleLine.Text | `14px`                          | Font size for button text.                                    |
| FontWeight              | Enum            | `SemiBold` (600)                | Font weight: 100–900.                                         |
| FontStyle               | Enum            | `Normal`                        | Font style: Normal, Italic, Oblique.                          |
| TextDecoration          | Enum            | `None`                          | Text decoration: None, Underline, Overline, Line-through.     |
| ButtonWidth             | SingleLine.Text | `auto`                          | CSS width (e.g., `150px`, `auto`).                            |
| ButtonHeight            | SingleLine.Text | `auto`                          | CSS height (e.g., `40px`, `auto`).                            |
| Padding                 | SingleLine.Text | `8px 16px`                      | Inner spacing.                                                |
| Margin                  | SingleLine.Text | `0px`                           | Outer spacing.                                                |
| BorderColor             | SingleLine.Text | `transparent`                   | CSS color for border.                                         |
| BorderWidth             | SingleLine.Text | `1px`                           | CSS width for border.                                         |
| BorderStyle             | Enum            | `Solid`                         | Border style: None, Hidden, Dotted, Dashed, Solid, etc.       |
| ButtonRadius            | SingleLine.Text | `4px`                           | CSS border-radius.                                            |
| ButtonBackgroundColor   | SingleLine.Text | `RGBA(0, 120, 212, 1)`          | Background color in normal state.                             |
| ButtonTextColor         | SingleLine.Text | `RGBA(255, 255, 255, 1)`        | Text color in normal state.                                   |
| HoverBackgroundColor    | SingleLine.Text | `RGBA(16, 110, 190, 1)`         | Background color on hover.                                    |
| HoverTextColor          | SingleLine.Text | `RGBA(255, 255, 255, 1)`        | Text color on hover.                                          |
| HoverBorderColor        | SingleLine.Text | `transparent`                   | Border color on hover.                                        |
| ActiveBackgroundColor   | SingleLine.Text | `RGBA(0, 90, 160, 1)`           | Background when active/clicked.                               |
| ActiveTextColor         | SingleLine.Text | `RGBA(255, 255, 255, 1)`        | Text color when active/clicked.                               |
| FocusBorderColor        | SingleLine.Text | `RGBA(0, 84, 153, 1)`           | Border color on focus.                                        |
| DisabledBackgroundColor | SingleLine.Text | `RGBA(243, 242, 241, 1)`        | Background color when disabled.                               |
| DisabledTextColor       | SingleLine.Text | `RGBA(161, 159, 157, 1)`        | Text color when disabled.                                     |
| ToolTip                 | SingleLine.Text | `Click to export data to Excel` | Tooltip text shown on hover.                                  |

## Usage Example

### Passing Static JSON Data

Use a static JSON string to represent your orders, including various data types:

```json
[
  { "OrderID": "1001", "OrderDate": "2025-05-01", "CustomerName": "Acme Corp", "OrderTotal": 1500.00, "IsPaid": true },
  { "OrderID": "1002", "OrderDate": "2025-05-03", "CustomerName": "Beta LLC", "OrderTotal": 875.50, "IsPaid": false }
]
```

Bind this directly in your component XML:

```xml
<property name="DataToExport" of-type="SingleLine.Text" default-value='[{"OrderID":"1001","OrderDate":"2025-05-01","CustomerName":"Acme Corp","OrderTotal":1500.00,"IsPaid":true},{"OrderID":"1002","OrderDate":"2025-05-03","CustomerName":"Beta LLC","OrderTotal":875.50,"IsPaid":false}]' />
```

### Exporting a Gallery's Data with Custom Controls

When your gallery contains nested controls, build the JSON manually. We use:

* **`Concat`**: iterates over each gallery item to build individual JSON object strings.
* **`Substitute`**: fixes the trailing comma issue by replacing `},]` with `}]`, yielding valid JSON.

Suppose you have:

* **Gallery**: `galOrders`
* **Controls**:

  * `lblOrderID` showing `ThisRecord.OrderID`
  * `lblOrderDate` showing `Text(ThisRecord.OrderDate, "yyyy-mm-dd")`
  * `lblCustomerName` showing `ThisRecord.CustomerName`
  * `lblOrderTotal` showing `Text(ThisRecord.OrderTotal, "0.00")`
  * `chkIsPaid` as a checkbox where `chkIsPaid.Value` is a boolean

Bind `DataToExport` to:

```powerapps
Substitute(
  "[" & Concat(
    galOrders.AllItems,
    "{"
    & "\"OrderID\": \"" & lblOrderID.Text & "\"," 
    & "\"OrderDate\": \"" & lblOrderDate.Text & "\"," 
    & "\"CustomerName\": \"" & lblCustomerName.Text & "\"," 
    & "\"OrderTotal\": " & lblOrderTotal.Text & "," 
    & "\"IsPaid\": " & If(chkIsPaid.Value, "true", "false") 
    & "},"
  ) & "]",
  "},]",
  "}]"
)
```

This results in the same JSON array as above:

```json
[
  { "OrderID": "1001", "OrderDate": "2025-05-01", "CustomerName": "Acme Corp", "OrderTotal": 1500.00, "IsPaid": true },
  { "OrderID": "1002", "OrderDate": "2025-05-03", "CustomerName": "Beta LLC", "OrderTotal": 875.50, "IsPaid": false }
]
```

### Exporting a Collection

Populate a collection with heterogeneous data types, then simply call `JSON`:

```powerapps
ClearCollect(
  colOrderData,
  { OrderID: "1001", OrderDate: Today() - 25, CustomerName: "Acme Corp", OrderTotal: 1500, IsPaid: true },
  { OrderID: "1002", OrderDate: Today() - 23, CustomerName: "Beta LLC", OrderTotal: 875.5, IsPaid: false }
);

// Bind DataToExport to:
JSON(colOrderData)
```

`JSON(colOrderData)` outputs exactly the same structured array, using property names as headers and properly formatting each value type.

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to open an issue or submit a pull request.

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE) for details.
