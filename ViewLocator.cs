using Avalonia.Controls;
using Avalonia.Controls.Templates;
using System;

namespace Pro7ChordEditor;

public class ViewLocator : IDataTemplate
{
    public Control? Build(object? data)
    {
        if (data is null)
            return null;

        // If it's a string or other simple type, just return a TextBlock
        if (data is string stringData)
        {
            return new TextBlock { Text = stringData };
        }

        // If it's a simple value type, convert to string
        if (data.GetType().IsValueType || data.GetType().IsPrimitive)
        {
            return new TextBlock { Text = data.ToString() };
        }

        var name = data.GetType().FullName!.Replace("ViewModel", "View", StringComparison.Ordinal);
        var type = Type.GetType(name);

        if (type != null)
        {
            var control = (Control)Activator.CreateInstance(type)!;
            control.DataContext = data;
            return control;
        }

        return new TextBlock { Text = data.ToString() };
    }

    public bool Match(object? data)
    {
        return data is not null;
    }
}