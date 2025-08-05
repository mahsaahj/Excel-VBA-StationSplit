# Excel VBA: Split Stations by Selected Parameters

## Overview
This Excel VBA project processes station data by splitting the dataset into separate worksheets based on unique station names.  
Users select desired parameters via a user form, and the code generates new sheets for each station with filtered data.  
Additionally, the average values of numeric columns are calculated and appended to each sheet.

## Features
- Dynamic reading of parameters from the header row  
- Multi-selection of parameters in a user-friendly form  
- Splitting data into separate worksheets for each station  
- Calculation and insertion of column-wise averages below the data  
- Handles columns for station name, year, and irrigation type (fixed columns) plus dynamic parameters  

## Usage Instructions
1. Open the Excel workbook (`.xlsm` format) with macros enabled.  
2. Run the macro `ShowParameterForm` to display the parameter selection form.  
3. Select one or more parameters to include in the output.  
4. Click **OK** to execute data splitting.  
5. Review the newly created worksheets, each named after a station, containing filtered data and average calculations.

## File Structure
- `SplitStationsByParameters` (Module): Core VBA subroutine for data splitting and average calculations.  
- `ParameterForm` (UserForm): Interface to select parameters dynamically from the dataset headers.  
- Main worksheet: Contains original raw data with station info and parameter values.

## Requirements & Notes
- The worksheet must have headers in the first row.  
- Station name must be in column D, year in column E, and irrigation type in column F.  
- Parameter columns should follow after column F.  
- Macros must be enabled in Excel to run the VBA code.  
- Existing worksheets named after stations will be deleted and replaced upon running the macro.  

## License
This project is licensed under the MIT License.

## Contact
For questions or feedback, please contact:  
your.email@example.com
