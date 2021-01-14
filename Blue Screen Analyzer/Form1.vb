Imports System.Diagnostics.Eventing.Reader
Imports System.Management
Imports System.Xml

Public Class Form1
    Dim OtherCred As Boolean = False
    Dim ObjFormCred As Net.NetworkCredential
    Dim UserName As String = Environment.UserName
    Dim UserDomain As String = Environment.UserDomainName
    Dim elSession As EventLogSession

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetWindowTitle()
        ObjFormCred = New Net.NetworkCredential(Environment.UserName, "", Environment.UserDomainName)
        UpdateCreds()
    End Sub

    Sub SetWindowTitle()
        ' Set the title of the form.
        Dim ApplicationTitle As String
        ApplicationTitle = String.Format("{0} - v{1}", My.Application.Info.ProductName, My.Application.Info.Version.ToString)
        Me.Text = ApplicationTitle
    End Sub


    Sub UpdateCreds()
        CredLabel.Text = ObjFormCred.Domain & "\" & ObjFormCred.UserName
    End Sub

    Function GetSecurePassword(ByVal StrPassword As String) As Security.SecureString
        Dim passwordChars As Char() = StrPassword.ToCharArray()
        Dim password As Security.SecureString = New Security.SecureString()
        For Each c As Char In passwordChars
            password.AppendChar(c)
        Next
        Return password
    End Function



    Function GetBugCheckEventLogs() As List(Of EventLogRecord)
        TextBox1.Text = ""

        If RemoteRadioButton.Checked = True And HostNameTextbox.Enabled Then
            TextBox1.Text = GetMachineInfo(True, HostNameTextbox.Text)
        Else
            TextBox1.Text = GetMachineInfo()
        End If

        Threading.Thread.Sleep(1000)


        Dim LogEntries As List(Of EventLogRecord) = New List(Of EventLogRecord)
        Dim elQuery As EventLogQuery = New EventLogQuery("System", PathType.LogName, "*[System/Provider/@EventSourceName='BugCheck']")

        ' /// Check If Remote Machine ///
        If RemoteRadioButton.Checked = True And HostNameTextbox.Enabled Then
            ' /// Remote Machine ... Setup Session ///
            If OtherCred Then
                elSession = New EventLogSession(HostNameTextbox.Text, ObjFormCred.Domain, ObjFormCred.UserName, GetSecurePassword(ObjFormCred.Password), SessionAuthentication.Default)
                elQuery.Session = elSession
            Else
                elSession = New EventLogSession(HostNameTextbox.Text)
                elQuery.Session = elSession

            End If



        End If

        Dim elReader As EventLogReader = New EventLogReader(elQuery)
        Try
            DataGridView1.Rows.Clear()
            Do Until 0 = 1
                Dim elEvent As EventLogRecord = elReader.ReadEvent()

                Dim EventID As String = elEvent.Id.ToString()
                Dim EventTimeStamp As String = elEvent.TimeCreated.ToString()
                Dim BugCheckCode As String = elEvent.Properties()(0).Value
                Dim DumpLocation As String = elEvent.Properties()(1).Value

                DataGridView1.Rows.Add(EventTimeStamp, BugCheckCode, DumpLocation)
            Loop
        Catch ex As Exception
            Exit Try
        End Try
        Return LogEntries
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetBugCheckEventLogs()
    End Sub

    Private Sub LocalRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles LocalRadioButton.CheckedChanged
        HostNameTextbox.Text = ""
        HostNameTextbox.Enabled = False
    End Sub

    Private Sub RemoteRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RemoteRadioButton.CheckedChanged
        HostNameTextbox.Text = ""
        HostNameTextbox.Enabled = True
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        StopCodeDetailsTextbox.Text = ""
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim ObjAbout As New AboutBox1
        ObjAbout.ShowDialog()
    End Sub


    Function GetMachineInfo(Optional Remote As Boolean = False, Optional Machine As String = "") As String
        On Error Resume Next

        Dim StrComputerName As String = ""
        Dim StrOperatingSystem As String = ""
        Dim StrLoggedOnUserName As String = ""


        If Remote Then
            ' /// Create Scope to Remote Machine ///
            Dim WMIScope As New ManagementScope("\\" & Machine & "\root\cimv2")

            If OtherCred Then
                WMIScope.Options.Username = ObjFormCred.Domain & "\" & ObjFormCred.UserName
                WMIScope.Options.Password = ObjFormCred.Password
            End If

            WMIScope.Connect()

            If WMIScope.IsConnected Then

                Dim OSQuery As New ObjectQuery("SELECT * FROM Win32_OperatingSystem")
                Dim CSQuery As New ObjectQuery("SELECT * FROM Win32_ComputerSystem")

                Dim os_searcher As New ManagementObjectSearcher(WMIScope, OSQuery)
                For Each info As ManagementObject In os_searcher.Get()
                    StrOperatingSystem = info.Properties("Caption").Value.ToString().Trim()
                Next info

                Dim cs_searcher As New ManagementObjectSearcher(WMIScope, CSQuery)
                For Each info As ManagementObject In cs_searcher.Get()
                    StrComputerName = info.Properties("name").Value.ToString().Trim()
                    StrLoggedOnUserName = info.Properties("username").Value.ToString().Trim()
                Next info

            End If

        Else

            Dim os_searcher As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
            For Each info As ManagementObject In os_searcher.Get()
                StrOperatingSystem = info.Properties("Caption").Value.ToString().Trim()
            Next info

            Dim cs_searcher As New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
            For Each info As ManagementObject In cs_searcher.Get()
                StrComputerName = info.Properties("name").Value.ToString().Trim()
                StrLoggedOnUserName = info.Properties("username").Value.ToString().Trim()
            Next info

        End If

        Dim StrComputerInfo As String = ""
        StrComputerInfo = "Host Name: " & StrComputerName & vbCrLf
        StrComputerInfo = StrComputerInfo & "Operating System: " & StrOperatingSystem & vbCrLf
        StrComputerInfo = StrComputerInfo & "Logged On UserName: " & StrLoggedOnUserName & vbCrLf

        Return StrComputerInfo
    End Function

























    Function GetStopCodeName(BugCheckCode) As String
        Dim StopCodeName As String = ""
        Dim FormattedBugCheckCode As String = Trim(BugCheckCode)
        FormattedBugCheckCode = "0x" & UCase(Strings.Right(FormattedBugCheckCode, 8))
        Select Case (FormattedBugCheckCode)
            Case "0x00000001"
                StopCodeName = "APC_INDEX_MISMATCH"
            Case "0x00000002"
                StopCodeName = "DEVICE_QUEUE_NOT_BUSY"
            Case "0x00000003"
                StopCodeName = "INVALID_AFFINITY_SET"
            Case "0x00000004"
                StopCodeName = "INVALID_DATA_ACCESS_TRAP"
            Case "0x00000005"
                StopCodeName = "INVALID_PROCESS_ATTACH_ATTEMPT"
            Case "0x00000006"
                StopCodeName = "INVALID_PROCESS_DETACH_ATTEMPT"
            Case "0x00000007"
                StopCodeName = "INVALID_SOFTWARE_INTERRUPT"
            Case "0x00000008"
                StopCodeName = "IRQL_NOT_DISPATCH_LEVEL"
            Case "0x00000009"
                StopCodeName = "IRQL_NOT_GREATER_OR_EQUAL"
            Case "0x0000000A"
                StopCodeName = "IRQL_NOT_LESS_OR_EQUAL"
            Case "0x0000000B"
                StopCodeName = "NO_EXCEPTION_HANDLING_SUPPORT"
            Case "0x0000000C"
                StopCodeName = "MAXIMUM_WAIT_OBJECTS_EXCEEDED"
            Case "0x0000000D"
                StopCodeName = "MUTEX_LEVEL_NUMBER_VIOLATION"
            Case "0x0000000E"
                StopCodeName = "NO_USER_MODE_CONTEXT"
            Case "0x0000000F"
                StopCodeName = "SPIN_LOCK_ALREADY_OWNED"
            Case "0x00000010"
                StopCodeName = "SPIN_LOCK_NOT_OWNED"
            Case "0x00000011"
                StopCodeName = "THREAD_NOT_MUTEX_OWNER"
            Case "0x00000012"
                StopCodeName = "TRAP_CAUSE_UNKNOWN"
            Case "0x00000013"
                StopCodeName = "EMPTY_THREAD_REAPER_LIST"
            Case "0x00000014"
                StopCodeName = "CREATE_DELETE_LOCK_NOT_LOCKED"
            Case "0x00000015"
                StopCodeName = "LAST_CHANCE_CALLED_FROM_KMODE"
            Case "0x00000016"
                StopCodeName = "CID_HANDLE_CREATION"
            Case "0x00000017"
                StopCodeName = "CID_HANDLE_DELETION"
            Case "0x00000018"
                StopCodeName = "REFERENCE_BY_POINTER"
            Case "0x00000019"
                StopCodeName = "BAD_POOL_HEADER"
            Case "0x0000001A"
                StopCodeName = "MEMORY_MANAGEMENT"
            Case "0x0000001B"
                StopCodeName = "PFN_SHARE_COUNT"
            Case "0x0000001C"
                StopCodeName = "PFN_REFERENCE_COUNT"
            Case "0x0000001D"
                StopCodeName = "NO_SPIN_LOCK_AVAILABLE"
            Case "0x0000001E"
                StopCodeName = "KMODE_EXCEPTION_NOT_HANDLED"
            Case "0x0000001F"
                StopCodeName = "SHARED_RESOURCE_CONV_ERROR"
            Case "0x00000020"
                StopCodeName = "KERNEL_APC_PENDING_DURING_EXIT"
            Case "0x00000021"
                StopCodeName = "QUOTA_UNDERFLOW"
            Case "0x00000022"
                StopCodeName = "FILE_SYSTEM"
            Case "0x00000023"
                StopCodeName = "FAT_FILE_SYSTEM"
            Case "0x00000024"
                StopCodeName = "NTFS_FILE_SYSTEM"
            Case "0x00000025"
                StopCodeName = "NPFS_FILE_SYSTEM"
            Case "0x00000026"
                StopCodeName = "CDFS_FILE_SYSTEM"
            Case "0x00000027"
                StopCodeName = "RDR_FILE_SYSTEM"
            Case "0x00000028"
                StopCodeName = "CORRUPT_ACCESS_TOKEN"
            Case "0x00000029"
                StopCodeName = "SECURITY_SYSTEM"
            Case "0x0000002A"
                StopCodeName = "INCONSISTENT_IRP"
            Case "0x0000002B"
                StopCodeName = "PANIC_STACK_SWITCH"
            Case "0x0000002C"
                StopCodeName = "PORT_DRIVER_INTERNAL"
            Case "0x0000002D"
                StopCodeName = "SCSI_DISK_DRIVER_INTERNAL"
            Case "0x0000002E"
                StopCodeName = "DATA_BUS_ERROR"
            Case "0x0000002F"
                StopCodeName = "INSTRUCTION_BUS_ERROR"
            Case "0x00000030"
                StopCodeName = "SET_OF_INVALID_CONTEXT"
            Case "0x00000031"
                StopCodeName = "PHASE0_INITIALIZATION_FAILED"
            Case "0x00000032"
                StopCodeName = "PHASE1_INITIALIZATION_FAILED"
            Case "0x00000033"
                StopCodeName = "UNEXPECTED_INITIALIZATION_CALL"
            Case "0x00000034"
                StopCodeName = "CACHE_MANAGER"
            Case "0x00000035"
                StopCodeName = "NO_MORE_IRP_STACK_LOCATIONS"
            Case "0x00000036"
                StopCodeName = "DEVICE_REFERENCE_COUNT_NOT_ZERO"
            Case "0x00000037"
                StopCodeName = "FLOPPY_INTERNAL_ERROR"
            Case "0x00000038"
                StopCodeName = "SERIAL_DRIVER_INTERNAL"
            Case "0x00000039"
                StopCodeName = "SYSTEM_EXIT_OWNED_MUTEX"
            Case "0x0000003A"
                StopCodeName = "SYSTEM_UNWIND_PREVIOUS_USER"
            Case "0x0000003B"
                StopCodeName = "SYSTEM_SERVICE_EXCEPTION"
            Case "0x0000003C"
                StopCodeName = "INTERRUPT_UNWIND_ATTEMPTED"
            Case "0x0000003D"
                StopCodeName = "INTERRUPT_EXCEPTION_NOT_HANDLED"
            Case "0x0000003E"
                StopCodeName = "MULTIPROCESSOR_CONFIGURATION_NOT_SUPPORTED"
            Case "0x0000003F"
                StopCodeName = "NO_MORE_SYSTEM_PTES"
            Case "0x00000040"
                StopCodeName = "TARGET_MDL_TOO_SMALL"
            Case "0x00000041"
                StopCodeName = "MUST_SUCCEED_POOL_EMPTY"
            Case "0x00000042"
                StopCodeName = "ATDISK_DRIVER_INTERNAL"
            Case "0x00000043"
                StopCodeName = "NO_SUCH_PARTITION"
            Case "0x00000044"
                StopCodeName = "MULTIPLE_IRP_COMPLETE_REQUESTS"
            Case "0x00000045"
                StopCodeName = "INSUFFICIENT_SYSTEM_MAP_REGS"
            Case "0x00000046"
                StopCodeName = "DEREF_UNKNOWN_LOGON_SESSION"
            Case "0x00000047"
                StopCodeName = "REF_UNKNOWN_LOGON_SESSION"
            Case "0x00000048"
                StopCodeName = "CANCEL_STATE_IN_COMPLETED_IRP"
            Case "0x00000049"
                StopCodeName = "PAGE_FAULT_WITH_INTERRUPTS_OFF"
            Case "0x0000004A"
                StopCodeName = "IRQL_GT_ZERO_AT_SYSTEM_SERVICE"
            Case "0x0000004B"
                StopCodeName = "STREAMS_INTERNAL_ERROR"
            Case "0x0000004C"
                StopCodeName = "FATAL_UNHANDLED_HARD_ERROR"
            Case "0x0000004D"
                StopCodeName = "NO_PAGES_AVAILABLE"
            Case "0x0000004E"
                StopCodeName = "PFN_LIST_CORRUPT"
            Case "0x0000004F"
                StopCodeName = "NDIS_INTERNAL_ERROR"
            Case "0x00000050"
                StopCodeName = "PAGE_FAULT_IN_NONPAGED_AREA"
            Case "0x00000051"
                StopCodeName = "REGISTRY_ERROR"
            Case "0x00000052"
                StopCodeName = "MAILSLOT_FILE_SYSTEM"
            Case "0x00000053"
                StopCodeName = "NO_BOOT_DEVICE"
            Case "0x00000054"
                StopCodeName = "LM_SERVER_INTERNAL_ERROR"
            Case "0x00000055"
                StopCodeName = "DATA_COHERENCY_EXCEPTION"
            Case "0x00000056"
                StopCodeName = "INSTRUCTION_COHERENCY_EXCEPTION"
            Case "0x00000057"
                StopCodeName = "XNS_INTERNAL_ERROR"
            Case "0x00000058"
                StopCodeName = "FTDISK_INTERNAL_ERROR"
            Case "0x00000059"
                StopCodeName = "PINBALL_FILE_SYSTEM"
            Case "0x0000005A"
                StopCodeName = "CRITICAL_SERVICE_FAILED"
            Case "0x0000005B"
                StopCodeName = "SET_ENV_VAR_FAILED"
            Case "0x0000005C"
                StopCodeName = "HAL_INITIALIZATION_FAILED"
            Case "0x0000005D"
                StopCodeName = "UNSUPPORTED_PROCESSOR"
            Case "0x0000005E"
                StopCodeName = "OBJECT_INITIALIZATION_FAILED"
            Case "0x0000005F"
                StopCodeName = "SECURITY_INITIALIZATION_FAILED"
            Case "0x00000060"
                StopCodeName = "PROCESS_INITIALIZATION_FAILED"
            Case "0x00000061"
                StopCodeName = "HAL1_INITIALIZATION_FAILED"
            Case "0x00000062"
                StopCodeName = "OBJECT1_INITIALIZATION_FAILED"
            Case "0x00000063"
                StopCodeName = "SECURITY1_INITIALIZATION_FAILED"
            Case "0x00000064"
                StopCodeName = "SYMBOLIC_INITIALIZATION_FAILED"
            Case "0x00000065"
                StopCodeName = "MEMORY1_INITIALIZATION_FAILED"
            Case "0x00000066"
                StopCodeName = "CACHE_INITIALIZATION_FAILED"
            Case "0x00000067"
                StopCodeName = "CONFIG_INITIALIZATION_FAILED"
            Case "0x00000068"
                StopCodeName = "FILE_INITIALIZATION_FAILED"
            Case "0x00000069"
                StopCodeName = "IO1_INITIALIZATION_FAILED"
            Case "0x0000006A"
                StopCodeName = "LPC_INITIALIZATION_FAILED"
            Case "0x0000006B"
                StopCodeName = "PROCESS1_INITIALIZATION_FAILED"
            Case "0x0000006C"
                StopCodeName = "REFMON_INITIALIZATION_FAILED"
            Case "0x0000006D"
                StopCodeName = "SESSION1_INITIALIZATION_FAILED"
            Case "0x0000006E"
                StopCodeName = "SESSION2_INITIALIZATION_FAILED"
            Case "0x0000006F"
                StopCodeName = "SESSION3_INITIALIZATION_FAILED"
            Case "0x00000070"
                StopCodeName = "SESSION4_INITIALIZATION_FAILED"
            Case "0x00000071"
                StopCodeName = "SESSION5_INITIALIZATION_FAILED"
            Case "0x00000072"
                StopCodeName = "ASSIGN_DRIVE_LETTERS_FAILED"
            Case "0x00000073"
                StopCodeName = "CONFIG_LIST_FAILED"
            Case "0x00000074"
                StopCodeName = "BAD_SYSTEM_CONFIG_INFO"
            Case "0x00000075"
                StopCodeName = "CANNOT_WRITE_CONFIGURATION"
            Case "0x00000076"
                StopCodeName = "PROCESS_HAS_LOCKED_PAGES"
            Case "0x00000077"
                StopCodeName = "KERNEL_STACK_INPAGE_ERROR"
            Case "0x00000078"
                StopCodeName = "PHASE0_EXCEPTION"
            Case "0x00000079"
                StopCodeName = "MISMATCHED_HAL"
            Case "0x0000007A"
                StopCodeName = "KERNEL_DATA_INPAGE_ERROR"
            Case "0x0000007B"
                StopCodeName = "INACCESSIBLE_BOOT_DEVICE"
            Case "0x0000007C"
                StopCodeName = "BUGCODE_NDIS_DRIVER"
            Case "0x0000007D"
                StopCodeName = "INSTALL_MORE_MEMORY"
            Case "0x0000007E"
                StopCodeName = "SYSTEM_THREAD_EXCEPTION_NOT_HANDLED"
            Case "0x0000007F"
                StopCodeName = "UNEXPECTED_KERNEL_MODE_TRAP"
            Case "0x00000080"
                StopCodeName = "NMI_HARDWARE_FAILURE"
            Case "0x00000081"
                StopCodeName = "SPIN_LOCK_INIT_FAILURE"
            Case "0x00000082"
                StopCodeName = "DFS_FILE_SYSTEM"
            Case "0x00000085"
                StopCodeName = "SETUP_FAILURE"
            Case "0x0000008B"
                StopCodeName = "MBR_CHECKSUM_MISMATCH"
            Case "0x0000008E"
                StopCodeName = "KERNEL_MODE_EXCEPTION_NOT_HANDLED"
            Case "0x0000008F"
                StopCodeName = "PP0_INITIALIZATION_FAILED"
            Case "0x00000090"
                StopCodeName = "PP1_INITIALIZATION_FAILED"
            Case "0x00000092"
                StopCodeName = "UP_DRIVER_ON_MP_SYSTEM"
            Case "0x00000093"
                StopCodeName = "INVALID_KERNEL_HANDLE"
            Case "0x00000094"
                StopCodeName = "KERNEL_STACK_LOCKED_AT_EXIT"
            Case "0x00000096"
                StopCodeName = "INVALID_WORK_QUEUE_ITEM"
            Case "0x00000097"
                StopCodeName = "BOUND_IMAGE_UNSUPPORTED"
            Case "0x00000098"
                StopCodeName = "END_OF_NT_EVALUATION_PERIOD"
            Case "0x00000099"
                StopCodeName = "INVALID_REGION_OR_SEGMENT"
            Case "0x0000009A"
                StopCodeName = "SYSTEM_LICENSE_VIOLATION"
            Case "0x0000009B"
                StopCodeName = "UDFS_FILE_SYSTEM"
            Case "0x0000009C"
                StopCodeName = "MACHINE_CHECK_EXCEPTION"
            Case "0x0000009E"
                StopCodeName = "USER_MODE_HEALTH_MONITOR"
            Case "0x0000009F"
                StopCodeName = "DRIVER_POWER_STATE_FAILURE"
            Case "0x000000A0"
                StopCodeName = "INTERNAL_POWER_ERROR"
            Case "0x000000A1"
                StopCodeName = "PCI_BUS_DRIVER_INTERNAL"
            Case "0x000000A2"
                StopCodeName = "MEMORY_IMAGE_CORRUPT"
            Case "0x000000A3"
                StopCodeName = "ACPI_DRIVER_INTERNAL"
            Case "0x000000A4"
                StopCodeName = "CNSS_FILE_SYSTEM_FILTER"
            Case "0x000000A5"
                StopCodeName = "ACPI_BIOS_ERROR"
            Case "0x000000A7"
                StopCodeName = "BAD_EXHANDLE"
            Case "0x000000AB"
                StopCodeName = "SESSION_HAS_VALID_POOL_ON_EXIT"
            Case "0x000000AC"
                StopCodeName = "HAL_MEMORY_ALLOCATION"
            Case "0x000000AD"
                StopCodeName = "VIDEO_DRIVER_DEBUG_REPORT_REQUEST"
            Case "0x000000B1"
                StopCodeName = "BGI_DETECTED_VIOLATION"
            Case "0x000000B4"
                StopCodeName = "VIDEO_DRIVER_INIT_FAILURE"
            Case "0x000000B8"
                StopCodeName = "ATTEMPTED_SWITCH_FROM_DPC"
            Case "0x000000B9"
                StopCodeName = "CHIPSET_DETECTED_ERROR"
            Case "0x000000BA"
                StopCodeName = "SESSION_HAS_VALID_VIEWS_ON_EXIT"
            Case "0x000000BB"
                StopCodeName = "NETWORK_BOOT_INITIALIZATION_FAILED"
            Case "0x000000BC"
                StopCodeName = "NETWORK_BOOT_DUPLICATE_ADDRESS"
            Case "0x000000BD"
                StopCodeName = "INVALID_HIBERNATED_STATE"
            Case "0x000000BE"
                StopCodeName = "ATTEMPTED_WRITE_TO_READONLY_MEMORY"
            Case "0x000000BF"
                StopCodeName = "MUTEX_ALREADY_OWNED"
            Case "0x000000C1"
                StopCodeName = "SPECIAL_POOL_DETECTED_MEMORY_CORRUPTION"
            Case "0x000000C2"
                StopCodeName = "BAD_POOL_CALLER"
            Case "0x000000C4"
                StopCodeName = "DRIVER_VERIFIER_DETECTED_VIOLATION"
            Case "0x000000C5"
                StopCodeName = "DRIVER_CORRUPTED_EXPOOL"
            Case "0x000000C6"
                StopCodeName = "DRIVER_CAUGHT_MODIFYING_FREED_POOL"
            Case "0x000000C7"
                StopCodeName = "TIMER_OR_DPC_INVALID"
            Case "0x000000C8"
                StopCodeName = "IRQL_UNEXPECTED_VALUE"
            Case "0x000000C9"
                StopCodeName = "DRIVER_VERIFIER_IOMANAGER_VIOLATION"
            Case "0x000000CA"
                StopCodeName = "PNP_DETECTED_FATAL_ERROR"
            Case "0x000000CB"
                StopCodeName = "DRIVER_LEFT_LOCKED_PAGES_IN_PROCESS"
            Case "0x000000CC"
                StopCodeName = "PAGE_FAULT_IN_FREED_SPECIAL_POOL"
            Case "0x000000CD"
                StopCodeName = "PAGE_FAULT_BEYOND_END_OF_ALLOCATION"
            Case "0x000000CE"
                StopCodeName = "DRIVER_UNLOADED_WITHOUT_CANCELLING_PENDING_OPERATIONS"
            Case "0x000000CF"
                StopCodeName = "TERMINAL_SERVER_DRIVER_MADE_INCORRECT_MEMORY_REFERENCE"
            Case "0x000000D0"
                StopCodeName = "DRIVER_CORRUPTED_MMPOOL"
            Case "0x000000D1"
                StopCodeName = "DRIVER_IRQL_NOT_LESS_OR_EQUAL"
            Case "0x000000D2"
                StopCodeName = "BUGCODE_ID_DRIVER"
            Case "0x000000D3"
                StopCodeName = "DRIVER_PORTION_MUST_BE_NONPAGED"
            Case "0x000000D4"
                StopCodeName = "SYSTEM_SCAN_AT_RAISED_IRQL_CAUGHT_IMPROPER_DRIVER_UNLOAD"
            Case "0x000000D5"
                StopCodeName = "DRIVER_PAGE_FAULT_IN_FREED_SPECIAL_POOL"
            Case "0x000000D6"
                StopCodeName = "DRIVER_PAGE_FAULT_BEYOND_END_OF_ALLOCATION"
            Case "0x000000D7"
                StopCodeName = "DRIVER_UNMAPPING_INVALID_VIEW"
            Case "0x000000D8"
                StopCodeName = "DRIVER_USED_EXCESSIVE_PTES"
            Case "0x000000D9"
                StopCodeName = "LOCKED_PAGES_TRACKER_CORRUPTION"
            Case "0x000000DA"
                StopCodeName = "SYSTEM_PTE_MISUSE"
            Case "0x000000DB"
                StopCodeName = "DRIVER_CORRUPTED_SYSPTES"
            Case "0x000000DC"
                StopCodeName = "DRIVER_INVALID_STACK_ACCESS"
            Case "0x000000DE"
                StopCodeName = "POOL_CORRUPTION_IN_FILE_AREA"
            Case "0x000000DF"
                StopCodeName = "IMPERSONATING_WORKER_THREAD"
            Case "0x000000E0"
                StopCodeName = "ACPI_BIOS_FATAL_ERROR"
            Case "0x000000E1"
                StopCodeName = "WORKER_THREAD_RETURNED_AT_BAD_IRQL"
            Case "0x000000E2"
                StopCodeName = "MANUALLY_INITIATED_CRASH"
            Case "0x000000E3"
                StopCodeName = "RESOURCE_NOT_OWNED"
            Case "0x000000E4"
                StopCodeName = "WORKER_INVALID"
            Case "0x000000E6"
                StopCodeName = "DRIVER_VERIFIER_DMA_VIOLATION"
            Case "0x000000E7"
                StopCodeName = "INVALID_FLOATING_POINT_STATE"
            Case "0x000000E8"
                StopCodeName = "INVALID_CANCEL_OF_FILE_OPEN"
            Case "0x000000E9"
                StopCodeName = "ACTIVE_EX_WORKER_THREAD_TERMINATION"
            Case "0x000000EA"
                StopCodeName = "THREAD_STUCK_IN_DEVICE_DRIVER"
            Case "0x000000EB"
                StopCodeName = "DIRTY_MAPPED_PAGES_CONGESTION"
            Case "0x000000EC"
                StopCodeName = "SESSION_HAS_VALID_SPECIAL_POOL_ON_EXIT"
            Case "0x000000ED"
                StopCodeName = "UNMOUNTABLE_BOOT_VOLUME"
            Case "0x000000EF"
                StopCodeName = "CRITICAL_PROCESS_DIED"
            Case "0x000000F1"
                StopCodeName = "SCSI_VERIFIER_DETECTED_VIOLATION"
            Case "0x000000F2"
                StopCodeName = "HARDWARE_INTERRUPT_STORM"
            Case "0x000000F3"
                StopCodeName = "DISORDERLY_SHUTDOWN"
            Case "0x000000F4"
                StopCodeName = "CRITICAL_OBJECT_TERMINATION"
            Case "0x000000F5"
                StopCodeName = "FLTMGR_FILE_SYSTEM"
            Case "0x000000F6"
                StopCodeName = "PCI_VERIFIER_DETECTED_VIOLATION"
            Case "0x000000F7"
                StopCodeName = "DRIVER_OVERRAN_STACK_BUFFER"
            Case "0x000000F8"
                StopCodeName = "RAMDISK_BOOT_INITIALIZATION_FAILED"
            Case "0x000000F9"
                StopCodeName = "DRIVER_RETURNED_STATUS_REPARSE_FOR_VOLUME_OPEN"
            Case "0x000000FA"
                StopCodeName = "HTTP_DRIVER_CORRUPTED"
            Case "0x000000FC"
                StopCodeName = "ATTEMPTED_EXECUTE_OF_NOEXECUTE_MEMORY"
            Case "0x000000FD"
                StopCodeName = "DIRTY_NOWRITE_PAGES_CONGESTION"
            Case "0x000000FE"
                StopCodeName = "BUGCODE_USB_DRIVER"
            Case "0x000000FF"
                StopCodeName = "RESERVE_QUEUE_OVERFLOW"
            Case "0x00000100"
                StopCodeName = "LOADER_BLOCK_MISMATCH"
            Case "0x00000101"
                StopCodeName = "CLOCK_WATCHDOG_TIMEOUT"
            Case "0x00000102"
                StopCodeName = "DPC_WATCHDOG_TIMEOUT"
            Case "0x00000103"
                StopCodeName = "MUP_FILE_SYSTEM"
            Case "0x00000104"
                StopCodeName = "AGP_INVALID_ACCESS"
            Case "0x00000105"
                StopCodeName = "AGP_GART_CORRUPTION"
            Case "0x00000106"
                StopCodeName = "AGP_ILLEGALLY_REPROGRAMMED"
            Case "0x00000108"
                StopCodeName = "THIRD_PARTY_FILE_SYSTEM_FAILURE"
            Case "0x00000109"
                StopCodeName = "CRITICAL_STRUCTURE_CORRUPTION"
            Case "0x0000010A"
                StopCodeName = "APP_TAGGING_INITIALIZATION_FAILED"
            Case "0x0000010C"
                StopCodeName = "FSRTL_EXTRA_CREATE_PARAMETER_VIOLATION"
            Case "0x0000010D"
                StopCodeName = "WDF_VIOLATION"
            Case "0x0000010E"
                StopCodeName = "VIDEO_MEMORY_MANAGEMENT_INTERNAL"
            Case "0x0000010F"
                StopCodeName = "RESOURCE_MANAGER_EXCEPTION_NOT_HANDLED"
            Case "0x00000111"
                StopCodeName = "RECURSIVE_NMI"
            Case "0x00000112"
                StopCodeName = "MSRPC_STATE_VIOLATION"
            Case "0x00000113"
                StopCodeName = "VIDEO_DXGKRNL_FATAL_ERROR"
            Case "0x00000114"
                StopCodeName = "VIDEO_SHADOW_DRIVER_FATAL_ERROR"
            Case "0x00000115"
                StopCodeName = "AGP_INTERNAL"
            Case "0x00000116"
                StopCodeName = "VIDEO_TDR_ERROR"
            Case "0x00000117"
                StopCodeName = "VIDEO_TDR_TIMEOUT_DETECTED"
            Case "0x00000119"
                StopCodeName = "VIDEO_SCHEDULER_INTERNAL_ERROR"
            Case "0x0000011A"
                StopCodeName = "EM_INITIALIZATION_FAILURE"
            Case "0x0000011B"
                StopCodeName = "DRIVER_RETURNED_HOLDING_CANCEL_LOCK"
            Case "0x0000011C"
                StopCodeName = "ATTEMPTED_WRITE_TO_CM_PROTECTED_STORAGE"
            Case "0x0000011D"
                StopCodeName = "EVENT_TRACING_FATAL_ERROR"
            Case "0x0000011E"
                StopCodeName = "TOO_MANY_RECURSIVE_FAULTS"
            Case "0x0000011F"
                StopCodeName = "INVALID_DRIVER_HANDLE"
            Case "0x00000120"
                StopCodeName = "BITLOCKER_FATAL_ERROR"
            Case "0x00000121"
                StopCodeName = "DRIVER_VIOLATION"
            Case "0x00000122"
                StopCodeName = "WHEA_INTERNAL_ERROR"
            Case "0x00000123"
                StopCodeName = "CRYPTO_SELF_TEST_FAILURE"
            Case "0x00000124"
                StopCodeName = "WHEA_UNCORRECTABLE_ERROR"
            Case "0x00000125"
                StopCodeName = "NMR_INVALID_STATE"
            Case "0x00000126"
                StopCodeName = "NETIO_INVALID_POOL_CALLER"
            Case "0x00000127"
                StopCodeName = "PAGE_NOT_ZERO"
            Case "0x00000128"
                StopCodeName = "WORKER_THREAD_RETURNED_WITH_BAD_IO_PRIORITY"
            Case "0x00000129"
                StopCodeName = "WORKER_THREAD_RETURNED_WITH_BAD_PAGING_IO_PRIORITY"
            Case "0x0000012A"
                StopCodeName = "MUI_NO_VALID_SYSTEM_LANGUAGE"
            Case "0x0000012B"
                StopCodeName = "FAULTY_HARDWARE_CORRUPTED_PAGE"
            Case "0x0000012C"
                StopCodeName = "EXFAT_FILE_SYSTEM"
            Case "0x0000012D"
                StopCodeName = "VOLSNAP_OVERLAPPED_TABLE_ACCESS"
            Case "0x0000012E"
                StopCodeName = "INVALID_MDL_RANGE"
            Case "0x0000012F"
                StopCodeName = "VHD_BOOT_INITIALIZATION_FAILED"
            Case "0x00000130"
                StopCodeName = "DYNAMIC_ADD_PROCESSOR_MISMATCH"
            Case "0x00000131"
                StopCodeName = "INVALID_EXTENDED_PROCESSOR_STATE"
            Case "0x00000132"
                StopCodeName = "RESOURCE_OWNER_POINTER_INVALID"
            Case "0x00000133"
                StopCodeName = "DPC_WATCHDOG_VIOLATION"
            Case "0x00000134"
                StopCodeName = "DRIVE_EXTENDER"
            Case "0x00000135"
                StopCodeName = "REGISTRY_FILTER_DRIVER_EXCEPTION"
            Case "0x00000136"
                StopCodeName = "VHD_BOOT_HOST_VOLUME_NOT_ENOUGH_SPACE"
            Case "0x00000137"
                StopCodeName = "WIN32K_HANDLE_MANAGER"
            Case "0x00000138"
                StopCodeName = "GPIO_CONTROLLER_DRIVER_ERROR"
            Case "0x00000139"
                StopCodeName = "KERNEL_SECURITY_CHECK_FAILURE"
            Case "0x0000013A"
                StopCodeName = "KERNEL_MODE_HEAP_CORRUPTION"
            Case "0x0000013B"
                StopCodeName = "PASSIVE_INTERRUPT_ERROR"
            Case "0x0000013C"
                StopCodeName = "INVALID_IO_BOOST_STATE"
            Case "0x0000013D"
                StopCodeName = "CRITICAL_INITIALIZATION_FAILURE"
            Case "0x00000140"
                StopCodeName = "STORAGE_DEVICE_ABNORMALITY_DETECTED"
            Case "0x00000141"
                StopCodeName = "VIDEO_ENGINE_TIMEOUT_DETECTED"
            Case "0x00000142"
                StopCodeName = "VIDEO_TDR_APPLICATION_BLOCKED"
            Case "0x00000143"
                StopCodeName = "PROCESSOR_DRIVER_INTERNAL"
            Case "0x00000144"
                StopCodeName = "BUGCODE_USB3_DRIVER"
            Case "0x00000145"
                StopCodeName = "SECURE_BOOT_VIOLATION"
            Case "0x00000147"
                StopCodeName = "ABNORMAL_RESET_DETECTED"
            Case "0x00000149"
                StopCodeName = "REFS_FILE_SYSTEM"
            Case "0x0000014A"
                StopCodeName = "KERNEL_WMI_INTERNAL"
            Case "0x0000014B"
                StopCodeName = "SOC_SUBSYSTEM_FAILURE"
            Case "0x0000014C"
                StopCodeName = "FATAL_ABNORMAL_RESET_ERROR"
            Case "0x0000014D"
                StopCodeName = "EXCEPTION_SCOPE_INVALID"
            Case "0x0000014E"
                StopCodeName = "SOC_CRITICAL_DEVICE_REMOVED"
            Case "0x0000014F"
                StopCodeName = "PDC_WATCHDOG_TIMEOUT"
            Case "0x00000150"
                StopCodeName = "TCPIP_AOAC_NIC_ACTIVE_REFERENCE_LEAK"
            Case "0x00000151"
                StopCodeName = "UNSUPPORTED_INSTRUCTION_MODE"
            Case "0x00000152"
                StopCodeName = "INVALID_PUSH_LOCK_FLAGS"
            Case "0x00000153"
                StopCodeName = "KERNEL_LOCK_ENTRY_LEAKED_ON_THREAD_TERMINATION"
            Case "0x00000154"
                StopCodeName = "UNEXPECTED_STORE_EXCEPTION"
            Case "0x00000155"
                StopCodeName = "OS_DATA_TAMPERING"
            Case "0x00000156"
                StopCodeName = "WINSOCK_DETECTED_HUNG_CLOSESOCKET_LIVEDUMP"
            Case "0x00000157"
                StopCodeName = "KERNEL_THREAD_PRIORITY_FLOOR_VIOLATION"
            Case "0x00000158"
                StopCodeName = "ILLEGAL_IOMMU_PAGE_FAULT"
            Case "0x00000159"
                StopCodeName = "HAL_ILLEGAL_IOMMU_PAGE_FAULT"
            Case "0x0000015A"
                StopCodeName = "SDBUS_INTERNAL_ERROR"
            Case "0x0000015B"
                StopCodeName = "WORKER_THREAD_RETURNED_WITH_SYSTEM_PAGE_PRIORITY_ACTIVE"
            Case "0x0000015C"
                StopCodeName = "PDC_WATCHDOG_TIMEOUT_LIVEDUMP"
            Case "0x0000015F"
                StopCodeName = "CONNECTED_STANDBY_WATCHDOG_TIMEOUT_LIVEDUMP"
            Case "0x00000160"
                StopCodeName = "WIN32K_ATOMIC_CHECK_FAILURE"
            Case "0x00000161"
                StopCodeName = "LIVE_SYSTEM_DUMP"
            Case "0x00000162"
                StopCodeName = "KERNEL_AUTO_BOOST_INVALID_LOCK_RELEASE"
            Case "0x00000163"
                StopCodeName = "WORKER_THREAD_TEST_CONDITION"
            Case "0x00000164"
                StopCodeName = "WIN32K_CRITICAL_FAILURE"
            Case "0x0000016C"
                StopCodeName = "INVALID_RUNDOWN_PROTECTION_FLAGS"
            Case "0x0000016D"
                StopCodeName = "INVALID_SLOT_ALLOCATOR_FLAGS"
            Case "0x0000016E"
                StopCodeName = "ERESOURCE_INVALID_RELEASE"
            Case "0x00000175"
                StopCodeName = "PREVIOUS_FATAL_ABNORMAL_RESET_ERROR"
            Case "0x00000178"
                StopCodeName = "ELAM_DRIVER_DETECTED_FATAL_ERROR"
            Case "0x0000017B"
                StopCodeName = "PROFILER_CONFIGURATION_ILLEGAL"
            Case "0x00000187"
                StopCodeName = "VIDEO_DWMINIT_TIMEOUT_FALLBACK_BDD"
            Case "0x00000188"
                StopCodeName = "CLUSTER_CSVFS_LIVEDUMP"
            Case "0x00000189"
                StopCodeName = "BAD_OBJECT_HEADER"
            Case "0x0000018B"
                StopCodeName = "SECURE_KERNEL_ERROR"
            Case "0x0000018E"
                StopCodeName = "KERNEL_PARTITION_REFERENCE_VIOLATION"
            Case "0x00000190"
                StopCodeName = "WIN32K_CRITICAL_FAILURE_LIVEDUMP"
            Case "0x00000191"
                StopCodeName = "PF_DETECTED_CORRUPTION"
            Case "0x00000192"
                StopCodeName = "KERNEL_AUTO_BOOST_LOCK_ACQUISITION_WITH_RAISED_IRQL"
            Case "0x00000193"
                StopCodeName = "VIDEO_DXGKRNL_LIVEDUMP"
            Case "0x00000195"
                StopCodeName = "SMB_SERVER_LIVEDUMP"
            Case "0x00000196"
                StopCodeName = "LOADER_ROLLBACK_DETECTED"
            Case "0x00000197"
                StopCodeName = "WIN32K_SECURITY_FAILURE"
            Case "0x00000198"
                StopCodeName = "UFX_LIVEDUMP"
            Case "0x00000199"
                StopCodeName = "KERNEL_STORAGE_SLOT_IN_USE"
            Case "0x0000019A"
                StopCodeName = "WORKER_THREAD_RETURNED_WHILE_ATTACHED_TO_SILO"
            Case "0x0000019B"
                StopCodeName = "TTM_FATAL_ERROR"
            Case "0x0000019C"
                StopCodeName = "WIN32K_POWER_WATCHDOG_TIMEOUT"
            Case "0x0000019D"
                StopCodeName = "CLUSTER_SVHDX_LIVEDUMP"
            Case "0x000001A3"
                StopCodeName = "CALL_HAS_NOT_RETURNED_WATCHDOG_TIMEOUT_LIVEDUMP"
            Case "0x000001A4"
                StopCodeName = "DRIPS_SW_HW_DIVERGENCE_LIVEDUMP"
            Case "0x000001C4"
                StopCodeName = "DRIVER_VERIFIER_DETECTED_VIOLATION_LIVEDUMP"
            Case "0x000001C5"
                StopCodeName = "IO_THREADPOOL_DEADLOCK_LIVEDUMP"
            Case "0x000001C8"
                StopCodeName = "MANUALLY_INITIATED_POWER_BUTTON_HOLD"
            Case "0x000001CC"
                StopCodeName = "EXRESOURCE_TIMEOUT_LIVEDUMP"
            Case "0x000001CD"
                StopCodeName = "INVALID_CALLBACK_STACK_ADDRESS"
            Case "0x000001CE"
                StopCodeName = "INVALID_KERNEL_STACK_ADDRESS"
            Case "0x000001CF"
                StopCodeName = "HARDWARE_WATCHDOG_TIMEOUT"
            Case "0x000001D0"
                StopCodeName = "CPI_FIRMWARE_WATCHDOG_TIMEOUT"
            Case "0x000001D1"
                StopCodeName = "TELEMETRY_ASSERTS_LIVEDUMP"
            Case "0x000001D2"
                StopCodeName = "WORKER_THREAD_INVALID_STATE"
            Case "0x000001D3"
                StopCodeName = "WFP_INVALID_OPERATION"
            Case "0x000001D4"
                StopCodeName = "UCMUCSI_LIVEDUMP"
            Case "0x00000356"
                StopCodeName = "XBOX_ERACTRL_CS_TIMEOUT"
            Case "0x00000BFE"
                StopCodeName = "BC_BLUETOOTH_VERIFIER_FAULT"
            Case "0x00000BFF"
                StopCodeName = "BC_BTHMINI_VERIFIER_FAULT"
            Case "0x00020001"
                StopCodeName = "HYPERVISOR_ERROR"
            Case "0x1000007E"
                StopCodeName = "SYSTEM_THREAD_EXCEPTION_NOT_HANDLED_M"
            Case "0x1000007F"
                StopCodeName = "UNEXPECTED_KERNEL_MODE_TRAP_M"
            Case "0x1000008E"
                StopCodeName = "KERNEL_MODE_EXCEPTION_NOT_HANDLED_M"
            Case "0x100000EA"
                StopCodeName = "THREAD_STUCK_IN_DEVICE_DRIVER_M"
            Case "0x4000008A"
                StopCodeName = "THREAD_TERMINATE_HELD_MUTEX"
            Case "0xC0000218"
                StopCodeName = "STATUS_CANNOT_LOAD_REGISTRY_FILE"
            Case "0xC000021A"
                StopCodeName = "STATUS_SYSTEM_PROCESS_TERMINATED"
            Case "0xC0000221"
                StopCodeName = "STATUS_IMAGE_CHECKSUM_MISMATCH"
            Case "0xDEADDEAD"
                StopCodeName = "MANUALLY_INITIATED_CRASH1"
            Case Else
        End Select
        Return StopCodeName
    End Function

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        For Each Row As DataGridViewRow In DataGridView1.SelectedRows
            ' /// Get Stop Code ///
            Dim StrStopCodeName As String = ""
            Dim StrStopCode As String = Row.Cells(1).Value
            StrStopCode = Strings.Left(Trim(StrStopCode), 10)
            StrStopCodeName = GetStopCodeName(StrStopCode)
            StopCodeDetailsTextbox.Text = "Stop Code: " & StrStopCode & " (" & StrStopCodeName & ")" & vbCrLf
        Next
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles CredLinkLabel.LinkClicked
        Dim CredForm As Credentials = New Credentials
        Dim Result As DialogResult = CredForm.ShowDialog()

        If Result = DialogResult.OK Then
            OtherCred = True
            ObjFormCred = New Net.NetworkCredential(CredForm.ObjCred.UserName, CredForm.ObjCred.Password, CredForm.ObjCred.Domain)
            UpdateCreds()
        Else
            If Result = DialogResult.No Then
                OtherCred = False
                ObjFormCred = New Net.NetworkCredential(Environment.UserName, "", Environment.UserDomainName)
                UpdateCreds()
            End If
        End If
    End Sub
End Class
