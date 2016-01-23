using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Common.Library
{
  /// <summary>
  /// A class for storing validation rule failure messages
  /// </summary>
  public class ValidationMessage
  {
    #region Constructors
    public ValidationMessage()
    {
    }

    public ValidationMessage(string message)
    {
      Message = message;
    }
    #endregion

    /// <summary>
    /// Get/Set the validation message to display
    /// </summary>
    public string Message { get; set; }
  }
}
