

# Advanced Product Excel Add-in


This repository contains a .NET Framework 4.8 solution for an advanced Excel Add-in. The solution includes:
- **ExcelIntegration**: An Excel Add-in (using Excel-DNA) that provides custom functions for advanced calculations and business logic integration.
- **BusinessLogic**: A business logic library with services and data transfer objects for calculations.
- **AddinLauncher**: A Windows Forms launcher for managing or testing business features.


The add-in extends Excel with domain-specific functions, leveraging backend services for enhanced workflows. The business logic is reusable across the add-in and launcher applications.

## Benefits
- **Seamless integration**: Advanced business logic is available directly within Excel, eliminating manual data transfer.
- **Custom functions**: Use domain-specific calculations and features as native Excel functions.
- **Reusable logic**: Shared backend logic across both the Excel add-in and the Windows Forms launcher.
- **Easy deployment**: Simple installation and usage for end users—no VBA or complex setup required.
- **Extends Excel**: Enhance Excel’s capabilities for your business needs without modifying spreadsheets or relying on macros.

## Projects
- **ExcelIntegration**: Excel Add-in using Excel-DNA
- **BusinessLogic**: Business logic and services
- **AddinLauncher**: Windows Forms launcher

## Getting Started
1. Clone the repository.
2. Open the solution in Visual Studio.
3. Restore NuGet packages.
4. Build the solution.
5. To use the Excel Add-in, build `ExcelIntegration` and follow instructions in `ExcelIntegration-AddIn.dna`.

## Requirements
- Visual Studio 2019 or later
- .NET Framework 4.8


## How to Use the Excel Add-in

1. **Build the Add-in**
   - Open the solution in Visual Studio.
   - Build the `ExcelIntegration` project. This will generate:
     - An `.xll` file (e.g., `ExcelIntegration-AddIn64-packed.xll`)—this is the file you load into Excel.
     - A `.dll` file (e.g., `ExcelIntegration.dll`)—this contains your add-in logic and is required by the `.xll` file.
   - Both files will be in the output directory (typically `bin/Debug` or `bin/Release`).

2. **Add the Add-in to Excel**
   - Open Microsoft Excel.
   - Go to the `File` menu, select `Options`, then `Add-ins`.
   - At the bottom, select `Excel Add-ins` and click `Go...`.
   - In the Add-ins dialog, click `Browse...` and navigate to the folder containing the built `.xll` file (e.g., `ExcelIntegration-AddIn64-packed.xll`).
   - Select the `.xll` file and click `OK` to load the add-in. (The `.dll` file must be in the same folder as the `.xll` file.)

3. **Use the Add-in Functions**
   - The custom functions provided by the add-in will now be available in Excel. You can use them like regular Excel functions.

## Usage
- The Excel Add-in exposes functions for use in Excel.
- Business logic is in `BusinessLogic`.

## Contributing

Pull requests are welcome. For major changes, please open an issue first.

## License
Licensing for this project is managed via the Windows registry. No LICENSE file is included in this repository. Please refer to the documentation or contact the author for details on license validation and usage.
