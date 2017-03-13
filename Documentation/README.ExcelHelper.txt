
In order to install the sample:

a. Compile .NET project ExcelHelper and place the resulting ExcelHelper.dll into {Innovator installation root}\Client\cbin
b. Create Innovator method (give it whatever name you like; let's call it here ExcelHelperActionMethod) with the code from the file ExcelHelperActionMethod.txt.
c. Create Innovator action (give it whatever name you like; let's call it here ExcelHelperAction) with the following properties:

Type: Generic
Location: Client
Label: Start File in Excel
Method: ExcelHelperActionMethod
Target: None

After that Innovator's main window will have menu Actions\Start File in Excel. When clicking on it an instance of Excel (MS Excel must be installed on your client machine) will pop up with the file specified in the ExcelHelperActionMethod.

