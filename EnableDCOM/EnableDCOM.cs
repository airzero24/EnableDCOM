using System;
using System.Management;
using System.ServiceProcess;
using System.Runtime.InteropServices;
using Win = Microsoft.Win32;
using static EnableDCOM.Pinvoke;

namespace EnableDCOM
{
    class EnableDCOM
    {
        public static bool CheckEnableDCOMRegistryValue(string ComputerName)
        {
            ManagementScope scope = null;
            string result = "";
            try
            {
                ConnectionOptions connection = new ConnectionOptions();
                connection.Impersonation = System.Management.ImpersonationLevel.Impersonate;

                scope = new ManagementScope($"\\\\{ComputerName}\\root\\default", connection);
                scope.Connect();

                ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);

                ManagementBaseObject inParams = registry.GetMethodParameters("GetStringValue");
                inParams["hDefKey"] = 2147483650;
                inParams["sSubKeyName"] = @"SOFTWARE\Microsoft\Ole";
                inParams["sValueName"] = @"EnableDCOM";
                ManagementBaseObject outParams = registry.InvokeMethod("GetStringValue", inParams, null);
                result = (string)outParams["sValue"];
                if (result == "Y")
                {
                    Console.WriteLine($"[+] DCOM is enabled on host {ComputerName}");
                    return true;
                }
                else
                {
                    Console.WriteLine($"[-] DCOM is disabled on host {ComputerName}");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool CheckRemoteRegistryValue(string ComputerName)
        {

            ManagementScope scope = null;
            UInt32 result;
            try
            {
                ConnectionOptions connection = new ConnectionOptions();
                connection.Impersonation = System.Management.ImpersonationLevel.Impersonate;

                scope = new ManagementScope($"\\\\{ComputerName}\\root\\default", connection);
                scope.Connect();

                ManagementClass registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);

                ManagementBaseObject inParams = registry.GetMethodParameters("GetDWORDValue");
                inParams["hDefKey"] = 2147483650;
                inParams["sSubKeyName"] = @"SYSTEM\CurrentControlSet\Services\RemoteRegistry";
                inParams["sValueName"] = @"Start";
                ManagementBaseObject outParams = registry.InvokeMethod("GetDWORDValue", inParams, null);
                result = (UInt32)outParams["uValue"];
                if (result == 3)
                {
                    Console.WriteLine($"[+] RemoteRegistry service is enabled on host {ComputerName}");
                    return true;
                }
                else
                {
                    Console.WriteLine($"[-] RemoteRegistry service is disabled on host {ComputerName}");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool EnableDcom(string ComputerName, string Value)
        {
            Win.RegistryKey baseKey = null;
            Win.RegistryKey DCOMKey = null;
            try
            {
                baseKey = Win.RegistryKey.OpenRemoteBaseKey(Win.RegistryHive.LocalMachine, ComputerName);
                Console.WriteLine("[+] Successfully got handle to HKLM hive on host " + ComputerName);
                DCOMKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Ole", true);
                Console.WriteLine("[+] Successfully opened DCOM key on host " + ComputerName);
                DCOMKey.SetValue("EnableDCOM", Value);
                Console.WriteLine("[+] Successfully created registry key value on host " + ComputerName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool ChangeServiceStartType(string ComputerName, string serviceName, ServiceStartupType startType)
        {
            try
            {
                IntPtr scmHandle = OpenSCManager(ComputerName, null, Pinvoke.SC_MANAGER_CONNECT);
                if (scmHandle == IntPtr.Zero)
                {
                    throw new Exception("[-] Failed to obtain a handle to the service control manager database.");
                }

                IntPtr serviceHandle = OpenService(scmHandle, serviceName, SERVICE_QUERY_CONFIG | SERVICE_CHANGE_CONFIG);
                if (serviceHandle == IntPtr.Zero)
                {
                    throw new Exception($"[-] Failed to obtain a handle to service '{serviceName}'.");
                }

                bool changeServiceSuccess = ChangeServiceConfig(serviceHandle, SERVICE_NO_CHANGE, (uint)startType, SERVICE_NO_CHANGE, null, null, IntPtr.Zero, null, null, null, null);
                if (!changeServiceSuccess)
                {
                    string msg = $"[-] Failed to update service configuration for service '{serviceName}'. ChangeServiceConfig returned error {Marshal.GetLastWin32Error()}.";
                    throw new Exception(msg);
                }

                if (scmHandle != IntPtr.Zero)
                    CloseServiceHandle(scmHandle);
                if (serviceHandle != IntPtr.Zero)
                    CloseServiceHandle(serviceHandle);
                Console.WriteLine("[+] Successfully updated RemoteRegistry service start type on host " + ComputerName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool StartRemoteRegistryService(string ComputerName, string serviceName)
        {
            try
            {
                IntPtr scmHandle = OpenSCManager(ComputerName, null, SC_MANAGER_CONNECT);
                if (scmHandle == IntPtr.Zero)
                {
                    throw new Exception("[-] Failed to obtain a handle to the service control manager database.");
                }

                IntPtr serviceHandle = OpenService(scmHandle, serviceName, SERVICE_START);
                if (serviceHandle == IntPtr.Zero)
                {
                    throw new Exception($"[-] Failed to obtain a handle to service '{serviceName}'.");
                }

                bool startSuccess = StartService(serviceHandle, 0, null);
                if (!startSuccess == true)
                {
                    string msg = $"[-] Failed to start service '{serviceName}'. Returned error {Marshal.GetLastWin32Error()}.";
                    throw new Exception(msg);
                }

                if (scmHandle != IntPtr.Zero)
                    CloseServiceHandle(scmHandle);
                if (serviceHandle != IntPtr.Zero)
                    CloseServiceHandle(serviceHandle);
                Console.WriteLine("[+] Successfully started RemoteRegistry service on host " + ComputerName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool StopRemoteRegistryService(string ComputerName, string serviceName)
        {
            try
            {
                var serviceInstance = new ServiceController(serviceName, ComputerName);
                serviceInstance.Stop();
                serviceInstance.WaitForStatus(ServiceControllerStatus.Stopped);
                Console.WriteLine("[+] Successfully stopped RemoteRegistry service on host " + ComputerName);
                serviceInstance.Close();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }

        public static bool CheckRunRegistryService(string ComputerName, string serviceName)
        {
            try
            {
                var serviceInstance = new ServiceController(serviceName, ComputerName);
                if (serviceInstance.Status == ServiceControllerStatus.Running)
                {
                    Console.WriteLine("[+] RemoteRegistry service is running on host " + ComputerName);
                    return true;
                }
                else if (serviceInstance.Status == ServiceControllerStatus.StartPending)
                {
                    Console.WriteLine("[+] RemoteRegistry service is starting on host " + ComputerName);
                    return true;
                }
                else
                {
                    Console.WriteLine($"[!] Remote Registry service is not started on host " + ComputerName);
                    return false;
                }
                serviceInstance.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
                return false;
            }
        }
    }
}
