using System;
using static EnableDCOM.EnableDCOM;
using static EnableDCOM.Pinvoke;

namespace EnableDCOM
{
    class Program
    {
        private static void Usage()
        {
            string Usage = @"
Usage:
    EnableDCOM.exe [option] [computername]
Parameters:
    option          Action to perform. Must be 'check', 'enable', or 'disable'. (Required)

                    check:
                        Check if DCOM is enabled on remote host.
                    enable:
                        Enable DCOM on remote host.
                    disable:
                        Disable DCOM on remote host.

    computername    Computer to perform the action against. (Required)

Example:
    Check if DCOM is enabled on remote host:
        EnableDCOM.exe check DC-01
    
    Enable DCOM on remote host:
        EnableDCOM.exe enable DC-01

    Disable DCOM on remote host:
        EnableDCOM.exe disable DC-01
";
            Console.WriteLine(Usage);
        }

        static void Main(string[] args)
        {
            try
            {
                switch(args[0].ToLower())
                {
                    case "check":
                        try
                        {
                            Console.WriteLine("[*] Checking if DCOM is enabled on host " + args[1]);
                            if (CheckEnableDCOMRegistryValue(args[1]))
                            {
                                Environment.Exit(1);
                            }
                            else
                            {
                                Environment.Exit(1);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"[-] Error: {e.Message}");
                            Environment.Exit(1);
                        }
                        break;
                    case "enable":
                        try
                        {
                            Console.WriteLine("[*] Checking if DCOM is enabled on host " + args[1]);
                            if (CheckEnableDCOMRegistryValue(args[1]))
                            {
                                Environment.Exit(1);
                            }
                            else
                            {
                                Console.WriteLine("[*] Attempting to enable RemoteRegistry service on host " + args[1]);
                                if (ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Manual))
                                {
                                    Console.WriteLine("[*] Checking if RemoteRegistry service is running on host " + args[1]);
                                    if (CheckRunRegistryService(args[1], "RemoteRegistry")) {
                                        Console.WriteLine("[*] Attempting to enable DCOM on host " + args[1]);
                                        if (EnableDcom(args[1], "Y"))
                                        {
                                            Console.WriteLine("[*] Attempting to stop RemoteRegistry service on host " + args[1]);
                                            StopRemoteRegistryService(args[1], "RemoteRegistry");
                                            Console.WriteLine("[*] Attempting to disable RemoteRegistry service on host " + args[1]);
                                            ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Disabled);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("[*] Attempting to start RemoteRegistry service on host " + args[1]);
                                        if (StartRemoteRegistryService(args[1], "RemoteRegistry"))
                                        {
                                            Console.WriteLine("[*] Attempting to enable DCOM on host " + args[1]);
                                            if (EnableDcom(args[1], "Y"))
                                            {
                                                Console.WriteLine("[*] Attempting to stop RemoteRegistry service on host " + args[1]);
                                                StopRemoteRegistryService(args[1], "RemoteRegistry");
                                                Console.WriteLine("[*] Attempting to disable RemoteRegistry service on host " + args[1]);
                                                ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Disabled);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"[-] Error: {e.Message}");
                            Environment.Exit(1);
                        }
                        break;
                    case "disable":
                        try
                        {
                            Console.WriteLine("[*] Checking if DCOM is enabled on host " + args[1]);
                            if (!CheckEnableDCOMRegistryValue(args[1]))
                            {
                                Environment.Exit(1);
                            }
                            else
                            {
                                Console.WriteLine("[*] Attempting to enable RemoteRegistry service on host " + args[1]);
                                if (ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Manual))
                                {
                                    Console.WriteLine("[*] Checking if RemoteRegistry service is running on host " + args[1]);
                                    if (CheckRunRegistryService(args[1], "RemoteRegistry"))
                                    {
                                        Console.WriteLine("[*] Attempting to enable DCOM on host " + args[1]);
                                        if (EnableDcom(args[1], "N"))
                                        {
                                            Console.WriteLine("[*] Attempting to stop RemoteRegistry service on host " + args[1]);
                                            StopRemoteRegistryService(args[1], "RemoteRegistry");
                                            Console.WriteLine("[*] Attempting to disable RemoteRegistry service on host " + args[1]);
                                            ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Disabled);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("[*] Attempting to start RemoteRegistry service on host " + args[1]);
                                        if (StartRemoteRegistryService(args[1], "RemoteRegistry"))
                                        {
                                            Console.WriteLine("[*] Attempting to enable DCOM on host " + args[1]);
                                            if (EnableDcom(args[1], "N"))
                                            {
                                                Console.WriteLine("[*] Attempting to stop RemoteRegistry service on host " + args[1]);
                                                StopRemoteRegistryService(args[1], "RemoteRegistry");
                                                Console.WriteLine("[*] Attempting to disable RemoteRegistry service on host " + args[1]);
                                                ChangeServiceStartType(args[1], "RemoteRegistry", ServiceStartupType.Disabled);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"[-] Error: {e.Message}");
                            Environment.Exit(1);
                        }
                        break;
                    default:
                        Usage();
                        break;
                }
            }

            catch (Exception e)
            {
                Console.WriteLine($"[-] Error: {e.Message}");
            }
        }
    }
}
