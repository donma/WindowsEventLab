using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
namespace WindowsEventLab
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //關閉驗證的Policy 
            PowerShellHelper.ExecuteString("Set-ExecutionPolicy Unrestricted");
            var res1 = PowerShellHelper.InvokePowerShellScript(AppDomain.CurrentDomain.BaseDirectory + "script1.ps1");

            Console.WriteLine(res1);

            File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "log.txt", res1.ToString());
            Console.ReadLine();
        }

    }
}
