using System.Runtime.InteropServices;

namespace BomWizard.Services;

internal static class ActiveComObject
{
    public static object Get(string progId)
    {
        ThrowIfFailed(CLSIDFromProgID(progId, out var classId), $"Could not find COM program id '{progId}'.");
        ThrowIfFailed(
            GetActiveObject(ref classId, IntPtr.Zero, out var unknown),
            $"Could not connect to a running '{progId}' instance. Open SolidWorks before updating.");

        try
        {
            return Marshal.GetObjectForIUnknown(unknown);
        }
        finally
        {
            if (unknown != IntPtr.Zero)
            {
                Marshal.Release(unknown);
            }
        }
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);

    private static void ThrowIfFailed(int hresult, string message)
    {
        if (hresult < 0)
        {
            throw new COMException(message, hresult);
        }
    }
}
