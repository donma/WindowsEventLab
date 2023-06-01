using System;
using System.Diagnostics;
using System.Collections.ObjectModel;   //Used for CollectionSystem.Collections.ObjectModel.Collection
using System.ComponentModel;
using System.Data;
using System.Management.Automation.Runspaces;
using System.Management.Automation;
using System.Text;

namespace WindowsEventLab
{
    public enum ResultType
    {
        PSObjectCollection = 0,
        DataTable = 1
    }

    public class PowerShellHelper
    {

        public PowerShellHelper()
        {
        }

        //This works but only returns standard output as text and not an object but will still work to invoke full-fledged scripts
        //Object invokedResults = PowerShellHelper.InvokePowerShellScript(@"C:\MyDir\TestPoShScript.ps1");
        public static Object InvokePowerShellScript(string scriptPath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            Process process = new Process();
            Object returnValue = null;

            startInfo.FileName = @"powershell.exe";
        
            startInfo.Arguments = (@"& 'PATH'").Replace("PATH", scriptPath);
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
      
            process = new Process { StartInfo = startInfo, EnableRaisingEvents = true };
            process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
            process.StartInfo.StandardErrorEncoding = Encoding.UTF8;
            process.OutputDataReceived += new DataReceivedEventHandler
            (
                delegate (object sender, DataReceivedEventArgs e)
                {
                    //For some e.Data always has an empty string
                    returnValue = e.Data;
                    //using (StreamReader output = process.StandardOutput)
                    //{
                    //    standardOutput = output.ReadToEnd();
                    //}
                }
            );
            process.Start();
            //process.BeginOutputReadLine();  //This is starts reading the return value by invoking OutputDataReceived event handler
            process.WaitForExit();

            Object standardOutput = process.StandardOutput.ReadToEnd();
            //Assert.IsTrue(output.Contains("StringToBeVerifiedInAUnitTest"));

            String errors = process.StandardError.ReadToEnd();
            //Assert.IsTrue(string.IsNullOrEmpty(errors));

            process.Close();

            //For some reason returnValue does not have the object type output
            //return returnValue;
            return standardOutput;
        }

        //IDictionary parameters = new Dictionary<String, String>();
        //parameters.AddUser("Identity", "My-AD-Group-Name");
        //Collection<Object> results = PowerShellHelper.Execute(textBoxCommand.Text);
        //DataTable dataTable = PowerShellHelper.ToDataTable(results);
        public static Collection<Object> ExecuteString(string command)  //, IDictionary parameters
        {
            Collection<Object> results = null;
            string error = "";
            var ss = InitialSessionState.CreateDefault();
            //ss.ImportPSModulesFromPath("C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\Modules\\ActiveDirectory\\ActiveDirectory.psd1");
            //ss.ImportPSModulesFromPath(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\dbatools\dbatools.psm1");
            //ss.ImportPSModule(new[] { "ActiveDirectory", "dbatools" });

            using (var ps = PowerShell.Create(ss))
            {

                //http://www.agilepointnxblog.com/powershell-error-there-is-no-runspace-available-to-run-scripts-in-this-thread/
                // Exception getting "Path": "There is no Runspace available to run scripts in this thread. You can provide one in the DefaultRunspace property of the System.Management.Automation.Runspaces.Runspace type. The script block you attempted to invoke was: $this.Mainmodule.FileName"
                //
                //ps.AddCommand("[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null").Invoke();
                //var rslt = ps.AddCommand("Import-Module").AddParameter("Name", "ActiveDirectory").Invoke();
                //rslt = ps.AddCommand("Import-Module").AddParameter("Name", "dbatools").Invoke();
                //rslt = ps.AddCommand("Import-Module").AddParameter("Name", "sqlps").Invoke();

                if (ps.HadErrors)
                {
                    System.Collections.ArrayList errorList = (System.Collections.ArrayList)ps.Runspace.SessionStateProxy.PSVariable.GetValue("Error");
                    error = string.Join("\n", errorList.ToArray());
                    throw new Exception(error);
                }

                ps.Commands.Clear();

                PSInvocationSettings settings = new PSInvocationSettings();
                settings.ErrorActionPreference = ActionPreference.Stop;


                //results = ps.AddCommand(command).AddParameters(parameters).Invoke<PSObject>();
                //results = ps.AddCommand(command).AddParameters(parameters).Invoke<Object>();
                results = ps.AddScript(command).Invoke<Object>();

                if (ps.HadErrors == true)
                {
                    System.Collections.ArrayList errorList = (System.Collections.ArrayList)ps.Runspace.SessionStateProxy.PSVariable.GetValue("Error");
                    error = string.Join("\n", errorList.ToArray());
                    throw new Exception(error);
                }

                foreach (var result in results)
                {
                    Debug.WriteLine(result.ToString());
                }
            }

            return results;
        }


   

    }
}