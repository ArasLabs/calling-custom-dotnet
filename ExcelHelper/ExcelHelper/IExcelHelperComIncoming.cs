using System;
using System.Runtime.InteropServices;

namespace Utils
{
	/// <summary>
	/// COM interface for <see cref="ExcelHelper"/>.
	/// </summary>
	[ComVisible(true)]
  [Guid("C741E64C-1F49-4675-82BC-DB7BCF47851F")]
  [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	public interface IExcelHelperComIncoming
	{
    /// <summary>
    /// Creates a new COM object Excel.Application.
    /// This is equivalent of calling 
    /// <code>
    /// var excelApp = new ActiveXObject("Excel.Application");
    /// </code>
    /// in JScript.
    /// </summary>
    /// <returns>A reference to created Excel.Application</returns>
    object CreateExcelApplication();
	}
}
