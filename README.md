# Data Deliverable Tool

A tool to process data deliverable files, comparing to content of a supplied workbook to add any missing content and update existing content.

## Table of contents

- [Setup & Requirements](#setup-&-requirements)

- [GUI and How To Use](#gui-and-how-to-use)
  
  * [Deliverable File](#deliverable-file)
  
  * [Workbook File](#workbook-file)

- [Troubleshooting](#troubleshooting)

- [Details](#details)
  
  - [Headers](#headers)
  
  - [Validation Sheets](#validation-sheets)
  
  - [Site Inventory Sheet](#site-inventory-sheet)
  
  - [Work Items](#work-items)
    
    - [Deficiency Data](#deficiency-data)
    
    - [Cost Data](#cost-data)

- [Changing the code](#changing-the-code)
  
  - [Externalized Strings](#externalized-strings)
  
  - [Externalized Values](#externalized-values)
  
  - [Java Code](#java-code)

- [In the GitHub](#in-the-github)

- [License](#license)

## Setup & Requirements

Java SE 17 is required to run this program. If you've used any of my previous tools, you'll already have it installed. If you don't have Java 17 or newer, you can download an installer for Temurin/OpenJDK 17 from [here](https://github.com/adoptium/temurin17-binaries/releases/download/jdk-17.0.8%2B7/OpenJDK17U-jdk_x64_windows_hotspot_17.0.8_7.msi). This is an open-source version of java. Once downloaded, you can run the installer by double-clicking, it will open a window guiding you through the installation. Leaving everything as the defaults and just clicking through the pages should work perfectly.

The program itself can be downloaded from the [GitHub](https://github.com/Jaden-Unruh/Data-Deliverable-Tool), it is a `.jar` file in the parent directory, called something like `data-deliverable-tool-1.0.x-jar-with-dependencies.jar`. Click the name of the file there, and click the dowload button (an arrow pointing downards towards a tray) in the top-right. The button will say "Download raw file" when you hover over it. You can rename the file to whatever you like after it's downloaded.

Once Temurin/Java 17 and the program `.jar` are installed, double click the `.jar` to run.

## GUI and How To Use

After double-clicking the `.jar`, a window titled "Data Deliverable Tool" will open. It will have two prompts, as described below:

1. `Select a Deliverable CA file: Select...`
   
   * Click on the select button to open a file prompt, navigate to and select a deliverable inspection spreadsheet (`Deliverable - CA-20xx-⋯.xlsx`). Note that this must be a `*.xlsx` file, rather than `*.xlsb` or any other spreadsheet filetype - see Troubleshooting for more. The contents of this spreadsheet should be as described [below](#deliverable-sheet).

2. `Select a Workbook file: Select...`
   
   * Click on the select button and, as above, select a workbook spreadsheet (`Workbook.xlsx`). Again, this must be a `*.xlsx` file, and should have contents as described [below](#workbook-sheet).

The other contents of the window are the `Close`, `Run`, `Open deliverable location` and `Help` buttons. Close and run are self-explanatory. Help opens a brief dialogue describing what I've written above, with a prompt to go to the Github page for this extended README. Open deliverable location is initially disabled, but enables once a deliverable file is selected. It opens the Windows file explorer to the parent folder of the selected deliverable file; this is also where the output deliverable will be placed.

### Deliverable File

This spreadsheet will have many sheets within it, as named below. Note that the sheets can follow either the old naming pattern or the new one, and if they use the old names they will be updated to the new schema by the program:

- Building Validation Report -> Building Validation

- Grounds Validation Data -> Grounds Validation

- Tower Validation Data -> Tower Validation

- Tank Validation Data -> Tank Validation

- Site Inventory -> Asset Validation

- Work Order List (O&M)

- Work Orders -> Work Order Validation (DM/UK)

- Deficiency Data -> New Work Orders

- Cost Data

- Cost Factors

If any sheets have names that aren't on the list, they won't be renamed and will not halt the program. However, sheet names *must* match the prompted names for the validation steps to work.

Details on what each of the required sheets for validation should contain and what will happen to the sheets when the program is run can be found in [Details](#details).

### Workbook File

This should have 3 pages in it, titled `BTG Validation`, `Site Inventory`, and `Work Items`. They can be in any order, but must be named exactly as specified. This spreadsheet will not be edited at all by the program. `BTG Validation` is a list of buildings with identifiers (location number, name, etc.) and information (size, floors, gps coordinates, etc.). `Site Inventory` is a list of assets with identifiers (asset id, maximo id, name, etc.) and information (install year, RSL, CRV, etc.). Work items is a list of work items with identifiers. Details on exact placement of this data can be found in [Details](#details)

## Troubleshooting

> Nothing's happening when I double click the `.JAR` file

Ensure you've installed Java as specified under [Setup](#setup-&-requirements). If you believe you have, try checking your java version:

1. Press Win+R, type `cmd` and press enter - this will open a command prompt window
2. Type `java -version` and press enter
3. If you've installed java as specified, the first line under your typing should read `openjdk version "17.0.8" 2023-07-18`[^2]. If, instead, it says `'java' is not recognized as an internal...` then java is not installed.

[^2]: If you had a version of java other than the one specified in Setup, this may show a different version, but should be similar. However, you probably wouldn't be in this troubleshooting step if this is the case.

---

> I only have spreadsheets of type `*.xlsb` or `*.csv` (or any other spreadsheet type) and the program won't open them

Open the spreadsheets in Microsoft Excel and select 'File -> Save As -> This PC' and choosing 'Excel Workbook (.xlsx)' from the drop-down. A full list of filetypes that Excel supports (and thus can be converted to .xlsx) can be found [here](https://learn.microsoft.com/en-us/deployoffice/compat/office-file-format-reference#file-formats-that-are-supported-in-excel).

---

> `Run` isn't doing anything

Ensure that you've selected two `*.xlsx` files. Spreadsheets of a different type will not work.

---

> I'm getting an error message popping up when I run the file

If you're getting an error message and you can't figure out what it's saying or how to fix it, reach out to me. If you click `More Info` on the error popup and copy the big text box, that text (a full stack trace on the error) can help me figure out what's going on.

---

> Something else is going wrong

Don't hesitate to reach out to me if you have any other issues - always happy to help.

## Details

There are a few main sections that the program runs: filling in the headers; completing the validation sheets, all of which are quite similar; completing the site inventory sheet; and completing the work item sheets.

### Headers

Immediately after renaming the sheets, the program will attempt to reorganize the columns of all of the sheets to match the prescribed order (found in `data-deliverable-tool-x.x.x.jar\dataDeliverableTool\columnHeaders.dat`, see [Externalized Strings](#externalized-strings) for more). Data should be maintained as long as the first cell in its column is one of the prescribed headers - this is case sensitive, so be careful. Any data not in such a column will be lost in the output file (not the input, of course, that file will remain unchanged).

### Validation Sheets

For the Building, Tower, Grounds, and Tank Validation sheets, the program will first pull the location number (AB######) from column D of each sheet, then compare that to column C from the Workbook Validation sheet. It will pull the relevant information, such as inspection date, CRV Value, Floors, GPS coordinates, etc.; depending on which validation sheet we're working on at the time, and copy that data back to the deliverable file.

Note, on the building validation sheet, that latitude and longitude values will be 'trimmed' - that is, they will never be copied with more than 11 characters (including the `-` and `.` if applicable), and they will be rounded to reach that point. For example, `-120.39475859` will be copied as `-120.394759`. The number of digits can be changed in `values.properties`, see [Externalized Values](#externalized-values).

### Site Inventory Sheet

For the site inventory sheet, things are a little more complicated. We start by pulling the Asset ID from column B of the Site Inventory sheet in the deliverable. Then, we do one of the following:

* If a corresponding maximo ID is found in column AY of the workbook, we copy relevant information (manufacturer, install date, etc.).

* If no corresponding maximo ID is found, or if one is found but it has description `Removed` (case insensitive), change Status (col F) to `DECOMMISSIONED`

Next, we look at any rows of the Workbook that haven't been used yet - i.e., any whose Maximo ID is not on the Deliverable. For each, we add a new row to the deliverable, copying some information (Inspection number, Site ID) from other rows, pulling some from the workbook (Priority, inspection date, etc.), and we prompt for the location ID.

### Work Items

The `Work Items` section of the workbook is used for the `Deficiency Data` (renamed to `New Work Orders`) and `Cost Data` sections of the deliverable:

#### Deficiency Data

This section will start empty, so we take each line of `Work Items` in the workbook and copy over data - Work Item Number, Location ID, Maximo ID, Work Item Name, Problem/Solution Statements. We take the Work Category and Rank from the same cell in the workbook, splitting them into two separate cells, and take a substring of the Distress Type to get Reason for Deficiency. We pull Inspection Number and Site ID from the first row of `Building Data`. Status is defaulted to "NEW" and IA Function to "F" - these values can be changed in `messages.properties` if you decompress the `.jar`.

#### Cost Data

First, we check every line of `Work Items` against what's already in `Cost Data`  (using Work Item Number) to see which ones we need to copy over. For those that aren't on `Cost Data` yet, we copy over relevant information - Work Item Number, Location ID, Burdened Total Cost, and Work Item Name. We pull Inspection Number and Site ID from the first row of `Building Data` again. Type and Line Type are defaulted to "MATERIAL" - this can be changed in `messages.properties` if you decompress the `.jar`.

## Changing the Code

The `.JAR` file is compiled and compressed, meaning all the code is not human-readable. You can decompress and recompress the file to change certain parts, like some of the GUI text and default values for the sheets (see [Externalized Strings](#externalized-strings)), but all of the code itself is not editable. Instead, all of the program files are included in a [github repository](https://github.com/Jaden-Unruh/Data-Deliverable-Tool) so that anyone other than me could download them and open them in an IDE (I use Eclipse).

### Externalized Strings

Nearly every user-visible piece of text, both in the GUI and preset values used in the sheets are in a (somewhat) user-editable file - that is, you can edit it without recompiling the java code. You do, though, need to decompress and recompress the contents of the `.jar` using a tool like [WinRar](https://www.win-rar.com/).

The externalized strings are found in `data-deliverable-tool-x.x.x.jar\dataDeliverableTool\messages.properties`. Each is one line, constructed as a key-value pair, where everything to the left of the '=' is the key, used by the program to find the String to use (don't change that side), and everything to the right can be edited to change what the program uses whenever it references that key. For example, take the line `Main.sheet.cost.lineType=MATERIAL`. This is what the program will put as the default value in the LINETYPE column of the `Cost Data` sheet. If you changed that line to `Main.sheet.cost.lineType=SomethingElse`, then the program would put "SomethingElse" in the LINETYPE column for each row it adds to `Cost Data`. Be sure to save and recompress the `.jar` before running.

A similar method can be used to change the map of sheet names (used for renaming sheets from the old standard), and changing the order and name of the columns in each sheet. This data is found in the files `data-deliverable-tool-x.x.x.jar\dataDeliverableTool\newNames.dat` and `data-deliverable-tool-x.x.x.jar\dataDeliverableTool\columnHeaders.dat`, respectively.

### Externalized Values

Many important numbers, notably every column number used by the program, is in another (somewhat) user-editable file. Again, like the externalized strings, you have to decompress the `.jar`.

**NOTE: Column numbers are zero-indexed** - this means that the first column, column A, has index 0, and B is 1.

These values are found in `data-deliverable-tool-x.x.x.jar\dataDeliverableTool\values.properties`. As above, each is one line constructed as a key-value pair. Take the line `colNum.delv.aval.inspDate=11`, for example. The key is `colNum.delv.aval.inspDate`, so we're defining a **col**umn **num**ber on the **del**i**v**erable workbook, on the **a**sset **val**idation sheet, and that is the **insp**ection **date**. That value is **11**, meaning we're looking at the 11-index column, which is column **L** (although this is the 12th column, it has index 11, since we start at 0).

If columns ever get moved around, these values can be easily changed without changing the code itself.

### Java Code

The actual code for the project, written in Java, cannot be edited without recompiling the project. Thus, I have provided all my project files in the GitHub repository. The code itself is located within [/src/main/java/dataDeliverableTool](https://github.com/Jaden-Unruh/Data-Deliverable-Tool/tree/master/src/main/java/dataDeliverableTool), with `Main.java` being the primary program file. To edit these, I would advise cloning the project from GitHub and opening in your preferred IDE, then re-building using a tool like [Maven](https://maven.apache.org/). I have included my `pom.xml` to facilitate the build.

## In the GitHub

The [github repository](https://github.com/Jaden-Unruh/Data-Deliverable-Tool) has a handful of files, but most of them are only necessary if you wish to modify the code.

The main `.jar` that you downloaded to run the project is just that - the project itself, all bundled up neatly and easy to use.

The two files titled README (`.md` and `.html`) are this long text document - the `.md` file is my preferred method for writing these sorts of things, but if you can't open that one the `.html` should do just fine, and opens in any browser. It's also rendered nicely below the file list in the github, so you don't have to download it.

`LICENSE` is the legal protections for this project, it's a strong copyleft license. See [License](#license) below for more.

`doc` is detailed documentation of my java code using [javadoc](https://en.wikipedia.org/wiki/Javadoc). You can download the folder and open `doc\index.html`, or, if you don't want to download it, [this link](https://html-preview.github.io/?url=https://raw.githubusercontent.com/Jaden-Unruh/Data-Deliverable-Tool/master/doc/dataDeliverableTool/package-summary.html) will take you to a tool that allows you to view the html files without downloading.

Everything else - `.settings`, `src`, `target`, `.classpath`, `.project`, and `pom.xml` are the project files, included so anyone can view and edit my code if desired. My advice would be to download the whole project (clone) and open it with an IDE - I use [Eclipse](https://eclipseide.org/).

## License

In my previous tools, I did not include a License, but for this one I decided to - primarily in case I'm not around to maintain the tool in the future. It shouldn't affect any use of the tool within Akana, and doesn't have any impact on the copyright of data edited by the code - only future distributions of the code itself.

Data Deliverable Tool is available under the [GNU General Public License v3.0](https://www.gnu.org/licenses/gpl-3.0.en.html) or later. In summary, this code is available to use, copy, and modify, under the condition that all derivative works contianing the code (not including sheets edited with the code) are released under the same license. This project is provided without liability or warranty. See the `LICENSE` file for more.
