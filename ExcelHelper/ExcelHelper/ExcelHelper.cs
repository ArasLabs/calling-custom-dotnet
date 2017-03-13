using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Utils
{
  /// <summary>
  /// Utils to work with MS Excel.
  /// </summary>
  [ComVisible(true)]
  [Guid("A94666AC-4070-4374-A813-57B11867EF16")]
  [ClassInterface(ClassInterfaceType.None)]
	public class ExcelHelper: IExcelHelperComIncoming
  {
    #region IExcelHelperComIncoming Members
    /// <summary>
    /// Creates a new COM object Excel.Application.
    /// This is equivalent of calling 
    /// <code>
    /// var excelApp = new ActiveXObject("Excel.Application");
    /// </code>
    /// in JScript.
    /// </summary>
    /// <returns>A reference to created Excel.Application</returns>
    public object CreateExcelApplication()
    {
      Type excelApplicationType = Type.GetTypeFromProgID("Excel.Application", true);
      object res = Activator.CreateInstance(excelApplicationType);

      return res;
    }

    #endregion
  }
}
