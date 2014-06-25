using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using Microsoft.Internal.AzureTest;

namespace AzureHelpers
{
    /// <summary>
    /// This class contains helper functions to run commands on remote VMs hosted on azure
    /// They use PowershellClient from Microsoft.Internal.AzureTest.RemotePowershell.dll
    /// System.Management.Automation.metadata should be in your GAC
    /// </summary>
    public static class AzureRemotePowershellHelpers
    {
        private const string systemDriveEnvString = "SystemDrive";
        private const string defaultSystemDrive = "C:";

        public static Collection<PSObject> GetRemoteComputerName(string machineDns, string username, string password,bool useSSL, bool skipCN, bool skipCA)
        {
            // connect to VM using Powershell
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);

            Console.WriteLine("Getting a remote connection using all defaults");
            RemoteConnection connection = psclient.GetDefaultPSConnection(useSSL, machineDns, credential, skipCN, skipCA);

            Console.WriteLine("Invoking command, getting remote computer name");
            Collection<PSObject> psResult = psclient.InvokeScript(connection, "gc env:computername");
            Console.WriteLine("success..");

            return psResult;
        }

        /// <summary>
        /// Run Script On Remote VM by providing local path of the script file
        /// </summary>
        /// <param name="machineDns">Use HostedServiceDns</param>
        /// <param name="username">Azure VM Username</param>
        /// <param name="password">Azure VM Password</param>
        /// <param name="useSSL">use SSL true:https or false:http</param>
        /// <param name="skipCN">Skip CN check</param>
        /// <param name="skipCA">Skip CA check</param>
        /// <param name="scriptFilePath">The local path to script that needs to be run</param>
        /// <param name="port">winrm port to connect to : port = useSSL ? PowershellClient.defaultHTTPSport : PowershellClient.defaultHTTPport if using default</param>
        /// <returns></returns>
        public static Collection<PSObject> RunScriptOnRemoteVM(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA, int port,string scriptFilePath)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection using all defaults");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);
            connection.connectionInfo.MaximumConnectionRedirectionCount = 5;
            StreamReader sr = new StreamReader(scriptFilePath);
            string psscript = sr.ReadToEnd();

            Collection<PSObject> psResult = psclient.InvokeScript(connection, psscript);
            return psResult;
        }

        /// <summary>
        /// Run Powershell command On Remote VM
        /// </summary>
        public static Collection<PSObject> RunRemoteCommand(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA, int port, string command)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection using all defaults");

            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);

            Collection<PSObject> psResult = psclient.InvokeScript(connection, command);
            return psResult;
        }

       public static Collection<PSObject> ReadRemoteRegistry(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA,  int port, string regPath, string regName)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection ...");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);

            Collection<CommandParameter> param = new Collection<CommandParameter>
                {
                    new CommandParameter("Path", regPath),  
                    new CommandParameter("Name", regName)
                };
            Console.WriteLine("Reading registry..");
            Collection<PSObject> psResult = psclient.InvokeCommand(connection, "Get-ItemProperty", param);
            return psResult;
        }

        public static Collection<PSObject> ReadRemoteFile(string machineDns, string username, string password, bool useSSL, bool skipCN,
                                                          bool skipCA, int port, string filePath, string encoding)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection ...");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);
            Console.WriteLine("Reading File..");

            Collection<PSObject> psResult = psclient.InvokeScript(connection, string.Format("gc -Encoding {0} {1}", encoding, filePath));
            return psResult;
        }

        public static Collection<PSObject> ReadRemoteEnvironmentVariable(string machineDns, string username, string password,
                                                                         bool useSSL, bool skipCN, bool skipCA, int port, string envVarName)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection ...");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);
            Console.WriteLine("Reading Environment variable..");

            Collection<PSObject> psResult = psclient.InvokeScript(connection, string.Format("gc env:{0}", envVarName));
            return psResult;
        }

        public static bool SysprepRemoteMachine(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA, int port)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection ...");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);
            connection.connectionInfo.OperationTimeout = 10 * 60 * 1000; //10 minutes for operation timeout

            var sysDir = GetSystemDriveOfRemoteVM(machineDns, username, password, true, true, true, port);

            string generatedFileName = Path.GetRandomFileName();
            string filename = string.Format("{0}.bat", generatedFileName.Substring(0, generatedFileName.IndexOf('.')));
            string filePath = Path.Combine(sysDir, filename);
            Console.WriteLine("filepath : {0}", filePath);

            string cmd = string.Format(@"{0}Windows\system32\sysprep\sysprep.exe /oobe /generalize /shutdown", sysDir);
            Console.WriteLine("cmd : {0}", cmd);

            WriteToFile(machineDns, username, password,true,true,true, port, filePath, cmd);

            try
            {
                psclient.InvokeScript(connection, filePath);
            }
            catch (System.Management.Automation.Remoting.PSRemotingTransportException)
            {
                Console.WriteLine("Caught PSRemotingTransportException .. Sleeping 2 minutes");
                Thread.Sleep(TimeSpan.FromMinutes(2));
                return true; //sysprep started (may not have completed - wait for rolestate stopped)
            }

            return false;
        }

        public static string GetSystemDriveOfRemoteVM(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA, int port)
        {
            var sysDir = ReadRemoteEnvironmentVariable(machineDns, username, password, useSSL, skipCN, skipCA, port, systemDriveEnvString).FirstOrDefault().ToString();
            Console.WriteLine("The system drive returned is {0} ", sysDir);
            if (string.IsNullOrWhiteSpace(sysDir))
            {
                sysDir = defaultSystemDrive;
            }
            sysDir = sysDir + @"\";

            return sysDir;
        }

        public static void WriteToFile(string machineDns, string username, string password, bool useSSL, bool skipCN, bool skipCA, int port, string filePath,string writeContent)
        {
            PowershellClient psclient = new PowershellClient();
            PSCredential credential = psclient.GetPSCredential(username, password);
            Console.WriteLine("Getting a remote connection ...");
            RemoteConnection connection = psclient.CreateRemoteConnection(useSSL, machineDns, port, PowershellClient.defaultappName, PowershellClient.defaultShellUri, credential, skipCN, skipCA);

            Collection<CommandParameter> param = new Collection<CommandParameter>();
            param.Add(new CommandParameter("path", filePath));
            param.Add(new CommandParameter("value", writeContent));
            psclient.InvokeCommand(connection, "Set-Content", param);
        }
    }
}
