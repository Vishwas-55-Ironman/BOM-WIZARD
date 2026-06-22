using System.Reflection;

namespace BomWizard.Services;

internal static class ComDispatch
{
    public static object? GetProperty(object target, string name, params object[] args)
    {
        return target.GetType().InvokeMember(
            name,
            BindingFlags.GetProperty,
            binder: null,
            target,
            args);
    }

    public static void SetProperty(object target, string name, params object[] args)
    {
        target.GetType().InvokeMember(
            name,
            BindingFlags.SetProperty,
            binder: null,
            target,
            args);
    }

    public static object? Invoke(object target, string name, params object[] args)
    {
        return target.GetType().InvokeMember(
            name,
            BindingFlags.InvokeMethod,
            binder: null,
            target,
            args);
    }

    public static int GetInt(object target, string propertyName)
    {
        return Convert.ToInt32(GetProperty(target, propertyName));
    }

    public static string GetString(object target, string propertyName, params object[] args)
    {
        return Convert.ToString(GetProperty(target, propertyName, args))?.Trim() ?? string.Empty;
    }
}
