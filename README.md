# Tableau-filters-extraction


Tableau workbooks are created as an XML file. Sometimes we need to document all the filters used in the tableau workbooks. It is literally not possible to open each worksheet in the workbook and capture these filters. This python code is used to parse the XML for filter configuration in each worksheet and store them in a csv. 

The same code can be used to parse the configuration XML and fetch formulas used in the calculation field. This will reduce lot of manual work required during documentation or workbook upgrades.

Note :- This is under development. Though not all filter functions are captured, most important filters are captured here. It can be extended to capture all. 
