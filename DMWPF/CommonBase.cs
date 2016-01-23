using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;

namespace DMWPF
{
  ///// <summary>
  ///// Class for All Entity & View Model classes to inherit from
  ///// Implements the INotifyPropertyChanged event
  ///// </summary>
  //public abstract class CommonBase : INotifyPropertyChanged
  //{
  //  #region Constructor
  //  protected CommonBase()
  //  {
  //  }
  //  #endregion

  //  #region INotifyPropertyChanged Implementation
  //  public event PropertyChangedEventHandler PropertyChanged;

  //  /// <summary>
  //  /// The PropertyChanged Event to raise to any UI object
  //  /// The event is only invoked if data binding is used
  //  /// </summary>
  //  /// <param name="propertyName">The property name that is changing</param>          
  //  protected virtual void RaisePropertyChanged(string propertyName)
  //  {
  //    PropertyChangedEventHandler handler = this.PropertyChanged;
  //    if (handler != null)
  //    {
  //      var e = new PropertyChangedEventArgs(propertyName);
  //      handler(this, e);
  //    }
  //  }
  //  #endregion
  //}


  public abstract class PropertyChangedBase : INotifyPropertyChanged
  {
    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>
    /// Helper method to set a property value, typically used in implementing a setter.
    /// Returns true if the property actually changed.
    /// </summary>
    protected bool SetProperty<T>(ref T backingField, T value, Expression<Func<T>> property)
    {
      var changed = !EqualityComparer<T>.Default.Equals(backingField, value);
      if (changed)
      {
        backingField = value;
        NotifyPropertyChanged<T>(property);
      }
      return changed;
    }

    /// <summary>
    /// Raises the PropertyChanged event for the specified property.
    /// </summary>
    public void NotifyPropertyChanged<T>(Expression<Func<T>> property)
    {
      if (PropertyChanged == null)
        return;

      var lambda = (LambdaExpression)property;

      MemberExpression memberExpression;
      if (lambda.Body is UnaryExpression)
      {
        var unaryExpression = (UnaryExpression)lambda.Body;
        memberExpression = (MemberExpression)unaryExpression.Operand;
      }
      else memberExpression = (MemberExpression)lambda.Body;

      PropertyChanged(this, new PropertyChangedEventArgs(memberExpression.Member.Name));
    }

    /// <summary>
    /// Raises the PropertyChanged event for the specified property.
    /// </summary>
    protected virtual void NotifyPropertyChanged(PropertyChangedEventArgs args)
    {
      if (PropertyChanged == null)
        return;

      PropertyChanged(this, args);
    }

  }

}