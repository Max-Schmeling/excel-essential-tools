# Excel Essential Tools
This is an addin for Excel, which extends the default functionality of Excel with more than 50 tools. The tools were developed to save the user time when doing monotonous tasks in Excel. All the tools I developed for the addin arise from a concrete need. Once the tool is installed, a new tab will appear in Excel as illustrated in the following screenshot.

<p align="center">
  <img width="100%" src="/screenshots/ribbon-tab.png">
</p>

I developed [another addin](https://github.com/Max-Schmeling/excel-contact-tools) for Excel, which can considerably reduce the time you spend in Excel, when working with contact information, such as when you're working in a human resource department.

# What can the addin do for you?
The addin is especially useful for users who work with large tables, because it features tools for navigating in worksheets quickly. It also features tools to manipulate multiple cells at once, such as removing whitespace characters, or adding a certain string to multiple cells at a specific position. One feature which I have been using regularly allows you to find duplicates in the columns of a table (see *Detect duplicates*), and shows you which columns are primary keys (100% unique items). Check out the screenshots, to see all of the available tools:
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
  <img width="100%" src="/screenshots/menu_text.png">
</p>
<p align="center">
  <img width="100%" src="/screenshots/menu_table.png">
</p>

# Compatibility

The add-in only works for local installations of Microsoft Excel, i.e., this add-in does not work with a browser-version of Microsoft Office.

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

# Disclaimer

Do whatever you want with the source code, but make it open source when you publish anything based on the contents of this repository! Also do not sue me and do not make me responsible for any damage caused by the use of the addin. When using VBA-based addins its always a good idea to save your work regularly.
