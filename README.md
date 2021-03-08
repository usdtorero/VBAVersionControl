# ExcelVersionControlTemplate
Allow GIT workflow to be done with Excel vba programming. The Template will export all vba modules into a subfolder (of Template file) /VisualBasic upon saving the file. It will also automatically import from this directory(if exists) upon opening the template

/VisualBasic/VersionControl.bas is the main code to import /export

/VisualBasic/ThisWorkbook.cls simply calls the main code upon saving or opening the excel file itself.

Inspired by User Demosthenex https://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
