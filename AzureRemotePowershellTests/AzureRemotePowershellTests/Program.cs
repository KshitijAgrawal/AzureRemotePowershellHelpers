using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AzureHelpers;


namespace AzureRemotePowershellTests
{
    class AzureRemotePowershellTests
    {
        static void Main(string[] args)
        {
            string machineDns = "YourHostedServiceName.cloudapp.net";
            string username = "Username";
            string password = "Password";
            int port = 5986; //or your public port of the WinRM endpoint on your IaaS VM

            var result = AzureRemotePowershellHelpers.GetRemoteComputerName(machineDns, username, password, true, true, true);
            Console.WriteLine(result.First().ToString());

            var result1 = AzureRemotePowershellHelpers.GetSystemDriveOfRemoteVM(machineDns, username, password, true, true, true, port);
            Console.WriteLine(result1.First().ToString());

            var result2 = AzureRemotePowershellHelpers.RunRemoteCommand(machineDns, username, password, true, true, true, port, @"Get-WmiObject -class Win32_ComputerSystemProduct -namespace root\CIMV2 | Select UUID");
            Console.WriteLine(result2.First().ToString());

            var result3 = AzureRemotePowershellHelpers.SysprepRemoteMachine(machineDns, username, password, true, true, true, port);
            Console.WriteLine(result3);
        }
    }
}
