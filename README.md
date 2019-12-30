# Enabled DCOM on remote host

This PoC enables DCOM on a remote host via the registry. A combination of WMI, the Service Control Manager, and .Net classes are used to enable or disable DCOM on the remote host. The order of operation is as follows:

1. Use WMI's `StdRegProv` Management Class to check the `HKLM\Software\Microsoft\Ole\EnableDCOM` registry key on a remote host to verify whether DCOM is enabled or not. 
2. If step one fails, then it's _most_ likely that DCOM is not enabled. Interaction with the remote host's Service Control Manager is used to verify if the `Remote Registry` service is enabled.
3. If the service is _not_ enabled, then the `Start Type` for the service is modified to enable it. If the service _is_ enabled, this step is skipped.
4. Check to see if the `Remote Registry` service is running. If not, then it is started.
5. Use WMI's `StdRegProv` again to enable/disable DCOM on the host by modifying the `HKLM\Software\Microsoft\Ole\EnableDCOM` registry key to either `Y` or `N`.
6. Restore the `Remote Registry` service to it's original configuration.

>Note: This method requires Administrative privileges as it modifies the Local Machine registry hive
>Note: Some PInvoke signatures were used due to backwards compatibility in .NET versions

## How to use
The project will need to be compile with Visual Studio. Can then be used from the Windows commandline or through some other method (such as Beacon's `execute-assembly` command).

```
Check if DCOM is enabled on remote host:
    EnableDCOM.exe check [computername]

Enable DCOM on remote host:
    EnableDCOM.exe enable [computername]

Disable DCOM on remote host:
    EnableDCOM.exe disable [computername]
```

## Detection
This technique interacts/modifies registry keys as well as the `Remote Registry` service.

Registry
- Monitor for modification of the `HKLM\Software\Microsoft\Ole\EnableDCOM` registry key value.
    - SACL's may be an effective means of alert generation for this. See: [Set-AuditRule](https://github.com/hunters-forge/Set-AuditRule)

Service Control Manager
- Monitor windows event logs
    - Windows Event Id `7040`
        - `Param1 = Remote Registry`
        - `Param3 = enabled`

wmiprvse.exe
- Process will spawn as a child of `wmiprvse.exe`. This will be noisy but can be used to correlate with registry writes or network RPC traffic from another host if data is available.

svchost.exe
- Process creation with commandline `C:\Windows\system32\svchost.exe -k localService -p -s RemoteRegistry` indicates the start of the `Remote Registry` service. This may be an indicator of suspicious activity if this service is normally disabled.

## Resources
- [DCOM](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-dcom/4a893f3d-bd29-48cd-9f43-d9777a4415b0)
- [.NET](https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registry?view=netframework-4.8)
- [StdRegProv](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/stdregprov)
- [Service Control Manager](https://docs.microsoft.com/en-us/windows/win32/services/service-control-manager)
- [PInvoke](https://www.pinvoke.net)
- [SharpSC](https://github.com/djhohnstein/SharpSC)
