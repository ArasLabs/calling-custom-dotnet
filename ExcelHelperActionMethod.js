var self = this;  
 
if (!window.myExcelHelper)  
{  
  function startMainWhenReady()  
  {  
    var srcElement = window.event.srcElement;  
    if (srcElement.readyState == 4)  
    {  
      srcElement.detachEvent("onreadystatechange", startMainWhenReady);  
      window.myExcelHelper = srcElement.object;  
      main.call(self);  
    }  
  }  
 
  // Replace assembly name and fully qualified class name on appropriate ones.
  var excelHelperClassId = top.aras.getBaseURL() + "/cbin/ExcelHelper.dll#Utils.ExcelHelper";
  // This method creates a link to the .NET custom control which is ExcelHelpler.dll
  // in this particulare sample.
  top.aras.uiAddConfigLink2Doc4Assembly(document, "ExcelHelper");  
  var tagObject = document.createElement("<OBJECT id='ExcelHelper' classId='" + excelHelperClassId + "'></OBJECT>");  
  tagObject.attachEvent("onreadystatechange", startMainWhenReady);  
 
  var headObject = document.getElementsByTagName("HEAD")[0];  
  headObject.appendChild(tagObject);  
}  
else  
{  
  main.call(self);  
}  
 
function main(self)  
{
  // Note that the context item that is passed to the method is not used here
  // (see section 5 of Innovator "Programmer's Guide" for more details on
  // context items). Instead for the purposes of the demo hardcoded folder and
  // file name are used (of course, replace them on appropriate ones). In general, 
  // client actions depend on the context item and get required info from it.
  var folder = "c:\\temp";  
  var fileName = "test.xls";  
 
  // Here is where we call a method from the custom .NET control
  // In this particular sample the method creates an instance of 
  // Excel application which is 2 lines later is used to open a workbook.
  // In general, an arbitrary .NET method could do a lot of things:
  // communicate to another process, read\write from\to file; etc.
  var xel = window.myExcelHelper.CreateExcelApplication();
  xel.visible = true;  
  var wb = xel.Workbooks.Open(folder + '\\' + fileName);  
} 