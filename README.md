# Two-Excel-Files-Compare-Project
## Introduction
ExcelCompare is an ASP.NET Core MVC application designed to compare and manipulate data from Excel files using the EPPlus library. It allows users to upload two Excel files, perform operations like merging and removing duplicates from one file, and compare data between the two files to identify single and missing records.

## Features
**ASP.NET Core MVC**: Web application framework
**EPPlus**: Library for reading and writing Excel files without Excel installed
**Excel Operations**: Merge and remove duplicates from Excel 1, compare Excel 1 with Excel 2 to find single and missing records

## Steps
**Clone the Repository**.
**Open Solution in Visual Studio**:
- Open the solution (ExcelCompare.sln) in Visual Studio.

**Restore Packages**:
- Ensure that all NuGet packages are restored for the solution.

**Build and Run the Application**:
- Build and run the application using Ctrl + F5.
## Usage:
- Upload Excel files using the provided UI.
- Select Excel 1 for merging and removing duplicate records.
- Select Excel 2 to compare with Excel 1 and generate two output files: one with all single records and another with missing records.
## Project Structure
**ExcelCompare**: Main ASP.NET Core MVC project.
**Controllers**: Contains the MVC controllers handling file uploads and operations.
**Models**: Contains the models for data representation.
**Views**: Contains the Razor views for user interface.
**wwwroot**: Contains static assets like CSS, JavaScript, and uploaded files.
- if file not download then go to wwwroot folder and add uploads folder.
- No database setup is required as the application does not use a database.

## Technologies Used
**ASP.NET Core MVC**: Web framework
**EPPlus**: Library for Excel operations

## Contact
For any questions or issues, please contact Vishal Jamwal at vishaljamwal402@gmail.com .