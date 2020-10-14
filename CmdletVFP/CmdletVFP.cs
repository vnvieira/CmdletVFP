using System;
using System.Management.Automation;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

[Cmdlet(VerbsCommon.Get, "VFP")]
public class VisualFoxProRun : Cmdlet
{
    [Parameter(Mandatory = true)]
    public string FileName { get; set; }

    [Parameter(Mandatory = false)]
    public string Options { get; set; }

    protected override void ProcessRecord()
    {
        object vfp, response;
        Type type;

        type = Type.GetTypeFromProgID("XVFP.XVFP");
        vfp = Activator.CreateInstance(type);

        object[] args = { Path.Combine(Directory.GetCurrentDirectory(), Path.ChangeExtension(FileName, ".prg")), Options };

        response = type.InvokeMember("run", BindingFlags.InvokeMethod, null, vfp, args);

        Marshal.ReleaseComObject(vfp);

        WriteObject(response);
    }
}
