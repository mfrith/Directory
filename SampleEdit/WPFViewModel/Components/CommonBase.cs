using System;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace Common.Library
{
  /// <summary>
  /// Class for All Entity & View Model classes to inherit from
  /// Implements the INotifyPropertyChanged event
  /// </summary>
  public abstract class CommonBase : INotifyPropertyChanged
  {
    #region Constructor
    protected CommonBase()
    {
    }
    #endregion

    #region INotifyPropertyChanged Implementation
    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>
    /// The PropertyChanged Event to raise to any UI object
    /// The event is only invoked if data binding is used
    /// </summary>
    /// <param name="propertyName">The property name that is changing</param>          
    protected virtual void RaisePropertyChanged(string propertyName)
    {
      PropertyChangedEventHandler handler = this.PropertyChanged;
      if (handler != null)
      {
        var e = new PropertyChangedEventArgs(propertyName);
        handler(this, e);
      }
    }
    #endregion
  }
}
