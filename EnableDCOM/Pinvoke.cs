using System;
using System.Runtime.InteropServices;

namespace EnableDCOM
{
    class Pinvoke
    {
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr OpenSCManager(string machineName, string databaseName, uint dwAccess);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr OpenService(IntPtr hSCManager, string lpServiceName, uint dwDesiredAccess);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool StartService(IntPtr hService, UInt32 dwNumServiceArgs, string lpServiceArgVectors);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool ControlService(IntPtr hService, UInt32 dwControl, ref SERVICE_STATUS lpServiceStatus);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool ChangeServiceConfig(
            IntPtr hService,
            uint nServiceType,
            uint nStartType,
            uint nErrorControl,
            string lpBinaryPathName,
            string lpLoadOrderGroup,
            IntPtr lpdwTagId,
            [In] char[] lpDependencies,
            string lpServiceStartName,
            string lpPassword,
            string lpDisplayName);

        [DllImport("advapi32.dll", EntryPoint = "CloseServiceHandle")]
        public static extern int CloseServiceHandle(IntPtr hSCObject);

        public const uint SC_MANAGER_CONNECT = 0x0001;
        public const uint SERVICE_QUERY_CONFIG = 0x00000001;
        public const uint SERVICE_CHANGE_CONFIG = 0x00000002;
        public const uint SERVICE_NO_CHANGE = 0xFFFFFFFF;
        public const uint SERVICE_START = 0x0010;

        public enum ServiceStartupType : uint
        {
            BootStart = 0,
            SystemStart = 1,
            Automatic = 2,
            Manual = 3,
            Disabled = 4
        }

        [Flags]
        public enum SERVICE_TYPES : int
        {
            SERVICE_KERNEL_DRIVER = 0x00000001,
            SERVICE_FILE_SYSTEM_DRIVER = 0x00000002,
            SERVICE_WIN32_OWN_PROCESS = 0x00000010,
            SERVICE_WIN32_SHARE_PROCESS = 0x00000020,
            SERVICE_INTERACTIVE_PROCESS = 0x00000100
        }

        public enum SERVICE_STATE : uint
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct SERVICE_STATUS
        {
            public static readonly int SizeOf = Marshal.SizeOf(typeof(SERVICE_STATUS));
            public SERVICE_TYPES dwServiceType;
            public SERVICE_STATE dwCurrentState;
            public uint dwControlsAccepted;
            public uint dwWin32ExitCode;
            public uint dwServiceSpecificExitCode;
            public uint dwCheckPoint;
            public uint dwWaitHint;
        }

        [Flags]
        public enum SERVICE_CONTROL : uint
        {
            STOP = 0x00000001,
            PAUSE = 0x00000002,
            CONTINUE = 0x00000003,
            INTERROGATE = 0x00000004,
            SHUTDOWN = 0x00000005,
            PARAMCHANGE = 0x00000006,
            NETBINDADD = 0x00000007,
            NETBINDREMOVE = 0x00000008,
            NETBINDENABLE = 0x00000009,
            NETBINDDISABLE = 0x0000000A,
            DEVICEEVENT = 0x0000000B,
            HARDWAREPROFILECHANGE = 0x0000000C,
            POWEREVENT = 0x0000000D,
            SESSIONCHANGE = 0x0000000E
        }
    }
}
