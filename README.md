# Excel Essential Tools
This is an addin for Excel, which extends the default functionality of Excel with more than 50 tools. The tools were developed to save the user time when doing monotonous tasks in Excel and some are meant to speed up Excel's default tools. It features tools for unexperienced and experienced users. A few tools overlap with those of Power-Query, but most of them solve problems, which Power-Query doesn't. The idea is to keep it easy to use for everybody. For that reason, there is also no documentation, because I tried to keep the tools self-explanatory. If you want to contribute to the development or show your appreciation, head to the section [**Contribution**](#Contribution). Once the addin is installed, a new tab will appear in Excel as illustrated in the following screenshot:

<p align="center">
  <img width="100%" src="/screenshots/ribbon-tab.png">
</p>

*I developed [another addin](https://github.com/Max-Schmeling/excel-contact-tools) for Excel, which can considerably reduce the time you spend in Excel, when working with contact information, such as when you're working in a human resource department.*

# What can the addin do for you?
Here is a quick summary of the features (from left to right in the screenshot).
* Group ***Navigation***: This group of tools is useful for users who work with large tables, because it allows you to navigate in worksheets quickly. Clicking one of the eight arrows will make the view jump to the last or first cell which contains data, depending on which direction you choose.
* Menu ***Select***: This menu contains tools to expand the current selection in various directions. That's also useful, if you work very large tables. To make an example, say you have a table with 500k rows, and you want to select the data of a column in that table (ie. not the entire column), whose cells are empty every other row. In this case, neither `CTRL`+`SPACE` nor `CTRL`+`SHIFT`+`UP/DOWN-ARROW` will help you - but the tool *Expand selection to the top/bottom* has got you covered.
* Menu ***Workbook***: This menu contains tools to manage the current workbook. This includes renaming, moving, deleting, or opening the directory of the current workbook in the file-explorer, without having to close or minimize the workbook. It also includes tools to hide or unhide multiple worksheets with one click.
* Menu ***Worksheet***: This menu contains tools to edit the current worksheet. This includes hiding and unhiding multiple columns at once, and resizing or inserting multiple cells at once. It also includes a tool to delete unused rows in a worksheet, in case you get hold of a workbook whose worksheets are endlessly scrollable.
* Menu ***Range***: This menu contains tools to edit a range of cells with one click. This includes tools to convert formulas to values, convert numbers stored as text to numbers (faster than Excel can do), apply a calculation to multiple number-cells, reset selected formats (this is useful when you have a worksheet which is headachingly bad formatted), and transforming ranges.
* Menu ***Text***: This menu contains tools to edit the text in multiple cells at once. You can remove excessive whitespace and invisible characters, remove specific text, remove characters at a specific position, and insert text at a specific position.
* Menu ***Table***: This menu contains tools to analyze and edit tables. You can find duplicates in each column of a table, find out which columns are unique (i.e. best for doing VLOOKUPs), split a table into multiple tables (also with separate worksheets or workbooks if you want to), combine multiple tables from multiple workbooks into one table, convert A1-references into table-references (e.g. [@[mycolumn]])

Check out the following screenshots or just give it a try:

<p align="center">
  <img width="100%" src="/screenshots/menu_select.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_workbook.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_worksheet.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_range.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_text.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_table.png">
</p>

# Compatibility

The addin only works for local installations of Microsoft Excel, i.e., this add-in does not work with a browser-version of Microsoft Office.

# Installation

1. Download the file named *excel-essential-tools-addin.xlam*
2. Install the addin-file according to the instructions for **xlam**-addins. Head to the [official Microsoft website](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) or simply ask google *How to install xlam-addins in Excel?*
4. Restart Excel
5. A new ribbon named *Contacts* should have appeared
6. Congrats! The addin is installed and you should be able to use it. If not, check out the troubleshooting-section.

# Troubleshooting

1. If a yellow banner appears at the top of your spreadsheet, you will have to click `Enable Content`, before the add-in becomes usable.
2. Head to *File >> Options >> Trust Center >> Macro Settings* in Excel and make sure that macros and the VBA-object model are both enabled (see screenshot):
<p align="center">
  <img width="100%" src="/screenshots/enable_macros2.png">
</p>

# Contribution

The addin is by no means perfect. If you want to contribute, there are two favorable ways:
1. If you find any imperfections such as typos, bugs, or performance issues, feel free to open an issue.
2. If you want to translate the addin into your language you can open an issue for that too. The basic multi-language functionality is already implemented, new languages just need to be added.
3. People told me they'd be willing to pay for some of the features. If you like what you use, I'd be happy to accept a dontation: [![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/donate?hosted_button_id=V84BCLCD8DQK4)

# Disclaimer

Do whatever you want with the source code, but make it open source when you publish anything based on the contents of this repository! Also do not sue me and do not make me responsible for any damage caused by the use of the addin. When using VBA-based addins its always a good idea to save your work regularly.
