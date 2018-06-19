Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class Form1
    ' PIPexport_v2, written in VS 2010
    ' This is a special version of PIPmanage to provide for
    '   scheduled unattended running on server
    ' Selected fields will be imported into SQL Server after PIP export
    ' PIP tables or imported into SQL Server 
    '       database:  [HealthInfo]..[tbl_PIPindex]
    '       replacing existing tables each time
    ' by Don Couvillion
    ' started Jan 26, 2016

    '---- Basic Variable List -------------
    Dim sqlstm As String = ""
    Dim sqlstm2 As String = ""
    Dim sqlstm3 As String = ""
    Dim errmsg As String = ""
    Dim msg01 As String = ""
    Dim msg02 As String = ""
    Dim msg03 As String = ""
    Dim Outline As String = ""
    Dim OL As String = ""
    Dim LoadComplete As Boolean
    Dim CurrentUser As String = ""
    Dim KeepLooping As Boolean = False
    Dim LoopLimit As Integer = 0
    Dim LoopPos As Integer = 0
    Dim AP As Int16 = 0
    Dim rc As Int32 = 0
    Dim rc2 As Int32 = 0
    Dim rc3 As Int32 = 0
    Dim rcb As Boolean = False
    Dim rb As Boolean = False
    Dim rf As Double = 0
    Dim rs As String = ""
    Dim rs1 As String = ""
    Dim rs2 As String = ""
    Dim rs3 As String = ""
    Dim rs4 As String = ""
    Dim rd As Date = Nothing
    Dim i As Int16 = 0
    Dim x As String = ""
    Dim xx As String = ""
    Dim xxx As String = ""
    Dim str01 As String = ""
    Dim str02 As String = ""
    Dim str03 As String = ""
    Dim str04 As String = ""
    Dim Fname As String = ""
    Dim Fpath As String = ""

    '---- Timer ----
    Dim WithEvents Timer1 As System.Windows.Forms.Timer
    Dim TimerCount As Int16 = 0
    Dim TimerDirection As String = "Up"
    Dim TimerLimitHigh As Integer = 10
    Dim TimerLimitLow As Integer = 0
    Dim MachineState As String = ""     '"Starting", "Processing", "Closing", "Finished", 

    Dim ArgName As String = ""
    Dim TableName As String = ""
    Dim ProcessStarted As Boolean = False
    Dim PaintCalls As Integer = 0
    Dim NoError As Boolean

    'If gWhichOne = "FFF" Then
    '    BackupPath = BackupBase & "FFF\MAIN.DBF"
    'ElseIf gWhichOne = "ELSH" Then
    '    BackupPath = BackupBase & "ELSH\MAIN.DBF"
    'Else
    '    BackupPath = ""
    'End If
    Dim BackupBase As String = "C:\Data\PIP\"
    Dim SourceBase As String = "G:\PIP\"
    Dim SourcePath As String = ""
    Dim BackupPath As String = ""
    Dim PIPtable As Object
    Dim PIProws As Integer = 0
    Dim PIProwsTotal As Integer = 0

    '---- Attached Classes ---- ---- ---- ---- ----
    Dim ErrorHandle1 As ErrorHandle = New ErrorHandle
    Dim PIPdb_FFF As PIPdb = New PIPdb("FFF", True)        'True for Local backup database, for this app
    Dim PIPdb_ELSH As PIPdb = New PIPdb("ELSH", True)      'True for Local backup database, for this app
    Dim PIPdb_SFF As PIPdb = New PIPdb("ELSH", True)
    Dim PIPdb_GH As PIPdb = New PIPdb("ELSH", True)

    Dim SQLdb1 As SQLdb = New SQLdb


    '==== ==== Form1_Load ==== ==== ==== Form1_Load ==== ==== ==== ==== ==== ==== ==== ==== ====
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '---- Init the timer ----
        Timer1 = New Timer()
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Timer1.Start()
        TimerCount = 5
        TimerLimitLow = 0
        TimerDirection = "Down"
        MachineState = "Starting"

        '---- Init a few things ----
        If Environment_IsDev() Then
            checkRefreshPIP.Checked = False
        Else
            checkRefreshPIP.Checked = True
        End If

        '---- Display the app status ----
        lblAppStatus.Text = "Form_Load"

        LoadComplete = False
        NoError = True
        PaintCalls = 0
        '---- Diagnostic ----
        'ListBox1.Items.Add("Form1_Load")
        Application.DoEvents()

        '---- Set the table to be exported/imported ----
        '   In future versions this could be set from a calling comand line argument
        TableName = "MAIN"

        '---- Display the environment: Development or Production
        rb = Environment_IsDev()

        '---- Main Process ---- See sub Timer1_Tick()
        ' See:  Export_ProcessMain() - - Started in sub Timer1_Tick()

    End Sub   'Form1_Load()
    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        lblAppStatus.Text = "Form1_Activated"

    End Sub   'Form1_Activated()
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        ' Runs after Form1_Activated, and every time screen changes in any way

        PaintCalls += 1
        '---- Diagnostic ----
        'ListBox1.Items.Add("Form1_Paint call : " & PaintCalls)
        'Application.DoEvents()

    End Sub   'Form1_Paint()

    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 
    '==== App main processes  ==== ==== ==== ==== ==== ==== ==== ==== ==== ====
    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 
    Private Sub Export_ProcessMain()

        '---- Diagnostic ----
        'ListBox1.Items.Add("Top of Export_ProcessMain()")

        MachineState = "Processing"

        lblAppStatus.Text = "Processing"
        Application.DoEvents()
        '---- Temp for dev ----
        'System.Threading.Thread.Sleep(5000)

        '==== Check connections, these are show stoppers ==== ==== ====
        '---- Database Connection to SQL Server ----
        NoError = SQLdb1.ConnectionGood
        If Not NoError Then
            ListBox1.Items.Add("Can't connect to SQL Server 162")
        End If
        '---- Database Connection to PIP.FFF ----
        If Environment_IsDev() Then
            PIPdb_FFF.InterActive = True
        Else
            PIPdb_FFF.InterActive = False
        End If

        If PIPdb_FFF.DBConnTest < 1 Then
            NoError = False
            ListBox1.Items.Add("Can't connect to C:\Data\PIP\FFF\ PIP 167")
            ListBox1.Items.Add("Process Stopped with errors 144")
        End If
        '---- Database Connection to PIP.ELSH ----
        If Environment_IsDev() Then
            PIPdb_ELSH.InterActive = True
        Else
            PIPdb_ELSH.InterActive = False
        End If

        If PIPdb_ELSH.DBConnTest < 1 Then
            NoError = False
            ListBox1.Items.Add("Can't connect to C:\Data\PIP\ELSH\ PIP 173")
            ListBox1.Items.Add("Process Stopped with errors 150")
        End If
        '---- Database Connection Display positive test results ----
        If NoError Then
            ListBox1.Items.Add("All database connections are good.")
        End If

        '==== Check Source & Destination folders ==== ==== ====
        If NoError Then
            NoError = FolderTest()
        End If

        '---- Diagnostic, stop here ----
        'NoError = False            ' be sure to comment out for production

        '==== The process ==== ==== ====
        If LoadComplete Then
            '---- Nothing to do, Avoid re-running this ----
        ElseIf NoError Then

            '---- It's the first run ----
            LoadComplete = True             'indicate complete

            ListBox1.Items.Add("Starting Process")
            Application.DoEvents()

            '==== Clear destination table for new input ==== ==== ====
            '---- Production / Development environment ----
            If Environment_IsDev() Then
                '---- We are in Development environment, choose if you want to clear SQL table ----
                '       for actual replacement of all the records
                NoError = Table_Clear()        'un-comment to clear
            Else
                '---- We are in Production environment, clear the SQL table ----
                '       for actual replacement of all the records
                NoError = Table_Clear()        'un-comment for production
            End If
        Else
            '---- In Error condition ----
            ListBox1.Items.Add("Process Stopped with errors 164")
        End If

        '==== Process one PIP table from multi PIP databases ==== ==== ====
        If NoError Then
            '---- Destination table cleared, add new entry ----

            '*********************************
            '**** Add PIP Databases below ****
            '**** such as "FFF", "ELSH" ******
            '*********************************

            '---- Process FFF & ELSH ----
            Export_ProcessPart("GRPHOME")
            Export_ProcessPart("SFF")
            Export_ProcessPart("FFF")
            Export_ProcessPart("ELSH")

        End If

        '---- Final Display ----
        ListBox1.Items.Add("")
        If NoError Then
            ListBox1.Items.Add("Total PIP records : " & PIProwsTotal)
            ListBox1.Items.Add("Export & Import Completed Successfully.")
        Else
            ListBox1.Items.Add("Process Stopped with errors 188")
        End If

        '---- Shut Down, Set timer & MachineState for closing ----
        '---- Production / Development environment ----
        If Environment_IsDev() Then   'checkRefreshPIP
            '---- In Development, leave app running ----
            'i = 4474
            ''---- Diagnostic -- Set the timer for shutdown ----
            'TimerCount = 10
            'TimerLimitLow = 0
            'TimerDirection = "Down"
            'MachineState = "Closing"
            'Timer1.Enabled = True

            MachineState = "Finished"
            lblAppStatus.Text = "Finished, X-out when you like."
        Else
            '---- In Production ----
            '---- Close this application on completion of it's task ----
            '---- because this environment is un-attended ----
            'CloseApp()
            '---- Set the timer for shutdown ----
            TimerCount = 20             'Seconds to shutdown
            TimerLimitLow = 0
            TimerDirection = "Down"
            MachineState = "Closing"
            lblAppStatus.Text = "Closing in a few seconds."
            Timer1.Enabled = True
        End If

    End Sub   'Export_ProcessMain()

    Private Sub Export_ProcessPart(ByVal gWhichOne As String)
        ' gWhichOne should be "FFF", "ELSH", Etc.
        ' DataSet in:  PIPdb1.dsTableDS.Tables("xTable")

        'ListBox1.Items.Add(">>>> top of Export_ProcessPart " & gWhichOne & ">>>>>")
        ListBox1.Items.Add(" - ")
        '---- Definitions from top of this class ----
        'Dim NoError As Boolean
        'Dim PIPdb_FFF As PIPdb = New PIPdb("FFF", False)
        'Dim PIPdb_ELSH As PIPdb = New PIPdb("ELSH", False)
        'Dim PIPtable As Object

        '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ====
        '==== Copy production PIP table(s) to a backup folder ==== ==== ==== ==== ==== ==== ====
        '====   Only certain the target table(s) are needed, i.e. "MAIN"
        '====   Other tables to consider:  "UNIT","BED_ACT","DXSN","SNDX1_3","SNDX4"
        '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ====
        If NoError Then
            '---- Production / Development environment ----
            If Environment_IsDev() Then
                '---- In Development, does user want to copy PIP files to backup? ----
                If checkRefreshPIP.Checked Then
                    '---- In Development, CheckBox is checked ----
                    NoError = FileCopy(gWhichOne)           'BackupPath is set in FileCopy()
                Else
                    '---- CheckBox is NOT checked, need to set BackupPath here ----
                    If gWhichOne = "FFF" Then
                        BackupPath = BackupBase & "FFF\MAIN.DBF"
                    ElseIf gWhichOne = "ELSH" Then
                        BackupPath = BackupBase & "ELSH\MAIN.DBF"
                    ElseIf gWhichOne = "SFF" Then
                        BackupPath = BackupBase & "SFF\MAIN.DBF"
                    ElseIf gWhichOne = "GRPHOME" Then
                        BackupPath = BackupBase & "GRPHOME\MAIN.DBF"
                    Else
                        BackupPath = ""
                    End If
                End If
              
            Else
                '---- In Production, copy files, ignore CheckBox ----
                NoError = FileCopy(gWhichOne)           'BackupPath is set in FileCopy()
            End If
        End If

        '==== Export the the PIP file in the backup folder ==== ==== ====
        If NoError Then
            '---- Get the given PIP table ----
            '   This is generic it can query for any table given in "BackupPath"
            '   Return in PIPdb1.dsTableDS.Tables("xTable")
            If gWhichOne = "FFF" Then
                PIProws = PIPdb_FFF.ExportTable(BackupPath)
                PIProwsTotal = PIProwsTotal + PIProws
                If PIProws = -1 Then
                    ListBox1.Items.Add("Table FFF." & TableName & " query error 187")
                    NoError = False
                ElseIf PIProws = 0 Then
                    ListBox1.Items.Add("Table FFF." & TableName & " warning no records 190")
                    NoError = False
                Else
                    ListBox1.Items.Add("Table FFF." & TableName & " records exported : " & PIProws)
                End If

                '---- PIP data now in this DataTable ----
                PIPtable = PIPdb_FFF.dsTableDS.Tables("xTable")

            ElseIf gWhichOne = "ELSH" Then
                PIProws = PIPdb_ELSH.ExportTable(BackupPath)
                PIProwsTotal = PIProwsTotal + PIProws
                If PIProws = -1 Then
                    ListBox1.Items.Add("Table ELSH." & TableName & " query error 200")
                    NoError = False
                ElseIf PIProws = 0 Then
                    ListBox1.Items.Add("Table ELSH." & TableName & " warning no records 203")
                    NoError = False
                Else
                    ListBox1.Items.Add("Table ELSH." & TableName & " records exported : " & PIProws)
                End If

                '---- PIP data now in this DataTable ----
                PIPtable = PIPdb_ELSH.dsTableDS.Tables("xTable")

            ElseIf gWhichOne = "SFF" Then
                PIProws = PIPdb_SFF.ExportTable(BackupPath)
                PIProwsTotal = PIProwsTotal + PIProws
                If PIProws = -1 Then
                    ListBox1.Items.Add("Table SFF." & TableName & " query error 378")
                    ListBox1.Items.Add("> " & BackupPath)

                    NoError = False
                ElseIf PIProws = 0 Then
                    ListBox1.Items.Add("Table SFF." & TableName & " warning no records 381")
                    NoError = False
                Else
                    ListBox1.Items.Add("Table SFF." & TableName & " records exported : " & PIProws)
                End If

                '---- PIP data now in this DataTable ----
                PIPtable = PIPdb_SFF.dsTableDS.Tables("xTable")

            ElseIf gWhichOne = "GRPHOME" Then
                PIProws = PIPdb_GH.ExportTable(BackupPath)
                PIProwsTotal = PIProwsTotal + PIProws
                If PIProws = -1 Then
                    ListBox1.Items.Add("Table SFF." & TableName & " query error 393")
                    NoError = False
                ElseIf PIProws = 0 Then
                    ListBox1.Items.Add("Table SFF." & TableName & " warning no records 396")
                    NoError = False
                Else
                    ListBox1.Items.Add("Table SFF." & TableName & " records exported : " & PIProws)
                End If

                '---- PIP data now in this DataTable ----
                PIPtable = PIPdb_GH.dsTableDS.Tables("xTable")

            End If
        End If

        '==== Import the data to SQL Server database ==== ==== ====
        If NoError Then
            If TableName = "MAIN" Then
                NoError = MAINtable_Import(gWhichOne)
            ElseIf TableName = "BED_ACT" Then
                'NoError = BedActtable_Import(gWhichOne)
                ListBox1.Items.Add("Import of BED_ACT table is not implemented 219")
            ElseIf TableName = "DXSN" Then
                'NoError = DXSNtable_Import(gWhichOne)
                ListBox1.Items.Add("Import of DXSN table is not implemented 222")
            ElseIf TableName = "SNDX1_3" Then
                'NoError = SNDX1_3table_Import(gWhichOne)
                ListBox1.Items.Add("Import of SNDX1_3 table is not implemented 225")
            ElseIf TableName = "SNDX4" Then
                'NoError = SNDX4table_Import(gWhichOne)
                ListBox1.Items.Add("Import of SNDX4 table is not implemented 228")

                '*************************
                '**** Add Tables Here ****
                '*************************
                'ElseIf TableName = "xxx" Then
            Else

            End If

        End If

    End Sub   'Export_ProcessPart()

    Private Function MAINtable_Import(ByVal PIPdiv As String) As Boolean
        ' Import the PIP "MAIN" table into SQL Server
        ' Data to import is in local DataTable:  "PIPtable"
        ' Return True is successful, False if not
        '
        Dim RowCount As Integer = 0
        Dim RowPos As Integer = 0
        Dim xValue As String = ""

        '==== Import data into SQL.MAIN table ==== ====
        ListBox1.Items.Add("Now Saving to SQL")
        Application.DoEvents()

        '---- Set up the loop ----
        KeepLooping = True
        RowCount = PIProws
        RowPos = 0
        '---- Main loop ----
        Do While KeepLooping            'And RowPos < 4
            '==== Copy data items to SQL Insert query ==== ====

            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("HOSP_NUM")) Then
                SQLdb1.pipHOSP_NUM = ""
            Else
                SQLdb1.pipHOSP_NUM = Trim(PIPtable.Rows(RowPos).Item("HOSP_NUM"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("FNAME")) Then
                SQLdb1.pipFNAME = ""
            Else
                SQLdb1.pipFNAME = Trim(PIPtable.Rows(RowPos).Item("FNAME"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("LNAME")) Then
                SQLdb1.pipLNAME = ""
            Else
                SQLdb1.pipLNAME = Trim(PIPtable.Rows(RowPos).Item("LNAME"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("MID_INIT")) Then
                SQLdb1.pipMID_INIT = ""
            Else
                SQLdb1.pipMID_INIT = Trim(PIPtable.Rows(RowPos).Item("MID_INIT"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("RACE")) Then
                SQLdb1.pipRACE = ""
            Else
                SQLdb1.pipRACE = Trim(PIPtable.Rows(RowPos).Item("RACE"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("ETHNIC")) Then
                SQLdb1.pipETHNIC = ""
            Else
                SQLdb1.pipETHNIC = Trim(PIPtable.Rows(RowPos).Item("ETHNIC"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("SEX")) Then
                SQLdb1.pipSEX = ""
            Else
                SQLdb1.pipSEX = Trim(PIPtable.Rows(RowPos).Item("SEX"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("BIRTHDATE")) Then
                SQLdb1.pipBIRTHDATE = "1/1/1900"
            Else
                SQLdb1.pipBIRTHDATE = PIPtable.Rows(RowPos).Item("BIRTHDATE")
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("SSN")) Then
                SQLdb1.pipSSN = ""
            Else
                SQLdb1.pipSSN = Trim(PIPtable.Rows(RowPos).Item("SSN"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("UNIT")) Then
                SQLdb1.pipUNIT = ""
            Else
                SQLdb1.pipUNIT = Trim(PIPtable.Rows(RowPos).Item("UNIT"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("STATUS")) Then
                SQLdb1.pipSTATUS = ""
            Else
                SQLdb1.pipSTATUS = Trim(PIPtable.Rows(RowPos).Item("STATUS"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("ADM_DATE")) Then
                SQLdb1.pipADM_DATE = "1/1/1900"
            Else
                rs = PIPtable.Rows(RowPos).Item("ADM_DATE")
                If rs = "12:00:00 AM" Then
                    rs = "1/1/1900"
                End If
                SQLdb1.pipADM_DATE = rs
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("ADM_TYPE")) Then
              
                SQLdb1.pipADM_TYPE = ""
            Else
                SQLdb1.pipADM_TYPE = Trim(PIPtable.Rows(RowPos).Item("ADM_TYPE"))
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("DC_DATE")) Then
                '---- For Discharge Date, when null, set it arbitrarily high ----
                SQLdb1.pipDC_DATE = "12/31/9999"                        'Means NOT discharged
            Else
                rs = PIPtable.Rows(RowPos).Item("DC_DATE")
                '---- Fix certain cases ----
                If rs = "12:00:00 AM" Then
                    '---- Substitute the Default date ----
                    rs = "12/31/9999"                                   'Means NOT discharged
                ElseIf InStr(rs, "/20") > 1 Then
                    '---- Normal date, nothing to change here ----
                ElseIf InStr(rs, "/19") > 1 Then
                    '---- Normal date, nothing to change here ----
                Else
                    '---- Unknown case, check it out in the "Output" pane ----
                    'Debug.WriteLine(rs)
                End If
                SQLdb1.pipDC_DATE = rs
            End If
            '---- 
            If IsDBNull(PIPtable.Rows(RowPos).Item("FACID")) Then
                'given:          PIPdiv/gWhichOne  gWhichOne = "SFF", "GRPHOME"
                If PIPdiv = "GRPHOME" Then
                    '---- for "GRPHOME" use "00000" ----
                    SQLdb1.pipFACID = "00000"
                    SQLdb1.pipDIV = "CHD"           'CHD means Community Homes Division
                Else
                    SQLdb1.pipFACID = ""
                End If
            ElseIf Trim(PIPtable.Rows(RowPos).Item("FACID")) = "" Then
                'given:          PIPdiv/gWhichOne  gWhichOne = "SFF", "GRPHOME"
                If PIPdiv = "GRPHOME" Then
                    '---- for "GRPHOME" use "00000" ----
                    SQLdb1.pipFACID = "00000"
                    SQLdb1.pipDIV = "CHD"           'CHD means Community Homes Division
                ElseIf PIPdiv = "SFF" Then
                    SQLdb1.pipFACID = "00334"
                    SQLdb1.pipDIV = "SFF"
                Else
                    SQLdb1.pipFACID = ""
                End If

            Else
                xValue = Trim(PIPtable.Rows(RowPos).Item("FACID"))
                SQLdb1.pipFACID = xValue
                '---- FACID ----
                '  00337        FFF
                '  00332        ELSH
                '  00334        SFF
                '  00000        CHD -- Community Homes Division
                If xValue = "00337" Then
                    SQLdb1.pipDIV = "FFF"
                ElseIf xValue = "00332" Then
                    SQLdb1.pipDIV = "ELSH"
                ElseIf PIPdiv = "GRPHOME" Then
                    SQLdb1.pipDIV = "CHD"
                ElseIf PIPdiv = "SFF" Then
                    SQLdb1.pipDIV = "SFF"
                Else
                    SQLdb1.pipDIV = ""
                End If
            End If
            '----
            If IsDBNull(PIPtable.Rows(RowPos).Item("LAST_DC")) Then
                SQLdb1.pipLAST_DC = "1/1/1900"
            Else
                SQLdb1.pipLAST_DC = PIPtable.Rows(RowPos).Item("LAST_DC")
            End If
            '---- 


            '---- Save one row of data ----
            rc = SQLdb1.MAIN_SaveOne()
            If rc = -1 Then
                ListBox1.Items.Add("Save error row:" & RowPos)
                Application.DoEvents()
                KeepLooping = False
            ElseIf rc = 0 Then
                ListBox1.Items.Add("Nothing saved row:" & RowPos)
                Application.DoEvents()
                KeepLooping = True
            Else
                '---- Normal, nothing to do here ----
            End If

            '---- Next ----
            RowPos += 1
            If RowPos = RowCount Then
                KeepLooping = False
            End If
            '---- Display progress ----
            'ListBox1.Items.Add(RowPos)
            'Application.DoEvents()
        Loop

        '---- Save Event Message ----
        If RowPos = RowCount Then
            '---- Loop completed normally ----
            Return True
        Else
            '---- Loop didn't finish with all records ----
            ErrorHandle1.ItemNote = "Abnormal End at: " & RowPos & " of " & RowCount
            ListBox1.Items.Add("Abnormal " & PIPdiv & " MAIN import")
            Application.DoEvents()
            Return False
        End If

    End Function   'MAINtable_Import()

    Private Function BedActtable_Import(ByVal PIPdiv As String) As Boolean
        ' Import the PIP "BED_ACT" table into SQL Server
        ' Return True is successful, False if not
        '**** Add code here to implement ****
        '**** See previous version "PIPexport_v1" for example ****
        Return True
    End Function   'BedActtable_Import

    Private Function DXSNtable_Import(ByVal PIPdiv As String) As Boolean
        ' Import the PIP "DXSN" table into SQL Server
        ' Return True is successful, False if not
        '**** Add code here to implement ****
        '**** See previous version "PIPexport_v1" for example ****
        Return True
    End Function   'DXSNtable_Import()

    Private Function SNDX1_3table_Import(ByVal PIPdiv As String) As Boolean
        ' Import the PIP "SNDX1_3" table into SQL Server
        ' Return True is successful, False if not
        '**** Add code here to implement ****
        '**** See previous version "PIPexport_v1" for example ****
        Return True
    End Function   'SNDX1_3table_Import()

    Private Function SNDX4table_Import(ByVal PIPdiv As String) As Boolean
        ' Import the PIP "SNDX4" table into SQL Server
        ' Return True is successful, False if not
        '**** Add code here to implement ****
        '**** See previous version "PIPexport_v1" for example ****
        Return True
    End Function   'SNDX4table_Import()


    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 
    '==== App service procedures   ==== ==== ==== ==== ==== ==== ==== ==== ====
    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 

    Private Function Table_Clear() As Boolean
        ' Clear / Delete all records in the final destination SQL table: 
        '       Database: HealthInfo    Table: tbl_PIPindex
        ' To be replaced by all new records
        ' Return True is successful, False if not

        '---- Clear the SQL Server table:  tbl_PIPindex ----
        ListBox1.Items.Add("---- Clearing SQL table ----")
        Application.DoEvents()
        rc = Me.SQLdb1.PIPindex_Clear
        If rc = -1 Then
            ListBox1.Items.Add("Query Error 88 - Clear table failed.")
            Application.DoEvents()
            '---- Failed to clear existing records, stop process ----
            ErrorHandle1.ItemDescripton = "Could NOT clear table 150"
            ErrorHandle1.ItemNote = "Trying to clear table : PIPindex"
            Return False
        ElseIf rc = 0 Then
            ListBox1.Items.Add("Warning NO old records cleared 154")
            Application.DoEvents()
            Return True
        Else
            ListBox1.Items.Add("Success - Cleared table tbl_PIPindex")
            Application.DoEvents()
            Return True
        End If

    End Function   'Table_Clear() 

    Private Function FileCopy(ByVal gWhichOne As String) As Boolean
        ' Copy a "MAIN" file 
        ' Return True is successful, False if not

        Dim sourceDir As String = "G:\PIP\" & gWhichOne & "\"
        Dim backupDir As String = BackupBase & gWhichOne & "\"
        Dim fName As String = TableName & ".DBF"                    'TableName set in Form1_Load()
        Dim FilePath As String = ""

        '---- Use the Path.Combine method to safely append the file name to the path.
        '---- Will overwrite if the destination file already exists.
        SourcePath = Path.Combine(sourceDir, fName)
        BackupPath = Path.Combine(backupDir, fName)

        If Not File.Exists(SourcePath) Then
            ListBox1.Items.Add("Problem: file not found : " & SourcePath)
            Application.DoEvents()
            '---- Failed to clear existing records, stop process ----
            ErrorHandle1.ItemDescripton = "Could NOT copy table"
            ErrorHandle1.ItemNote = "Trying to copy " & gWhichOne & ".MAIN"
            Return False
        Else
            Try
                File.Copy(SourcePath, BackupPath, True)        'True is for Overwrite
                ListBox1.Items.Add("File copied to : " & BackupPath)
                Application.DoEvents()
                Return True
            Catch copyError As IOException
                ListBox1.Items.Add("File copy error")
                Debug.WriteLine(copyError.Message)
                Return False
            End Try
        End If
    End Function   'FileCopy()

    Private Sub CloseApp()
        '**** Obsolete, see Timer1() below ****

        ListBox1.Items.Add("---- ---- ----")
        ListBox1.Items.Add("Closing app.")
        Application.DoEvents()

        lblAppStatus.Text = ""

        System.Threading.Thread.Sleep(5000)
        Me.Close()
    End Sub

    Private Function Environment_IsDev() As Boolean
        ' Is this app running in the development environment?
        ' Return True if it is
        ' Assume that dev environment is computer name:  HRPD8V1
        '                                  and user is:  DCouvill

        Dim envString As String
        Dim index As Integer = 1
        Dim MatchCount As Integer = 0

        envString = Environ(index)
        While envString <> ""
            'ListBox1.Items.Add(index & "   " & envString)
            'Debug.WriteLine(index & "   " & envString)

            If envString = "COMPUTERNAME=HRPD8V1" Then
                MatchCount += 1
            ElseIf envString = "USERNAME=DCouvill" Then
                MatchCount += 1
            End If
            '---- Next
            index += 1
            envString = Environ(index)
        End While

        If MatchCount = 2 Then
            lblEnvironment.Text = "Environment : Development"
            Return True
        Else
            lblEnvironment.Text = "Environment : Production"
            Return False
        End If

    End Function   'Environment_IsDev()

    Private Function FolderTest03(ByVal gFolder As String) As Boolean

        Dim currentUser As WindowsIdentity = WindowsIdentity.GetCurrent()

        For Each iRef As IdentityReference In currentUser.Groups
            Console.WriteLine(iRef.Translate(GetType(NTAccount)))
        Next

        '---- Results in Output pane ---- AD Groups !!!!
        'DHH\OMH-ELS-FS01 LevelSysAdmin
        'DHH\DHH-BEN-FS01 DATA
        'DHH\DHH-ISB-FS01 Data OMH
        'DHH\OMH-ELS-FS01 ITS Inventory
        Return True

    End Function
    Private Function FolderTest02(ByVal gFolder As String) As Boolean
        ' This worked
        Dim security As FileSecurity = File.GetAccessControl(gFolder)
        'Dim security As FileSecurity = Directory.GetAccessControl(gFolder)  'Doesn't work
        '                                                                                    System.Type
        Dim accessRules As AuthorizationRuleCollection = security.GetAccessRules(True, True, GetType(NTAccount))

        For Each rule As FileSystemAccessRule In accessRules
            Debug.WriteLine(rule.FileSystemRights, rule.IdentityReference.Value)
        Next
        '---- Results in Output pane ----
        'BUILTIN\Administrators: FullControl
        'NT AUTHORITY\SYSTEM: FullControl
        'BUILTIN\Users: ReadAndExecute, Synchronize
        'NT AUTHORITY\Authenticated Users: Modify, Synchronize

        Return True
    End Function

   
    Private Function FolderTest() As Boolean
        ' Does a folder exist and can it be written to?
        ' gFolder can be like either
        '   C:\Data\PIP
        '       or
        '   C:\Data\PIP\

        Dim xFolder As String = ""
        Dim xFilePath As String = ""
        Dim xNoError As Boolean = True

        '==== Check the data source folders ==== ==== ====
        '       We only need to read these
        xFolder = SourceBase & "FFF"
        xFilePath = xFolder & "\MAIN.DBF"
        If Directory.Exists(xFolder) Then
            If Not File.Exists(xFilePath) Then
                ListBox1.Items.Add("File NOT found: " & xFilePath)
                xNoError = False
            End If
        Else
            ListBox1.Items.Add("Folder NOT found: " & xFolder)
            xNoError = False
        End If
        '---- ----
        xFolder = SourceBase & "ELSH"
        xFilePath = xFolder & "\MAIN.DBF"
        If Directory.Exists(xFolder) Then
            If Not File.Exists(xFilePath) Then
                ListBox1.Items.Add("File NOT found: " & xFilePath)
                xNoError = False
            End If
        Else
            ListBox1.Items.Add("Folder NOT found: " & xFolder)
            xNoError = False
        End If


        '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ====
        '==== Check the data backup folders ==== ==== ==== ====
        '==== We need to read, write, & delete = ==== ==== ====
        '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ====

        '==== FFF ==== ==== FFF ==== ==== FFF ==== ==== FFF ==== ====
        xFolder = BackupBase & "FFF"
        xFilePath = xFolder & "\Test.txt"
        If Directory.Exists(xFolder) Then
            '==== Folder exist, can we write & delete? ==== ==== ====
            '---- File Delete ----
            Try
                If File.Exists(xFilePath) Then
                    Kill(xFilePath)
                End If
            Catch ex As Exception
                ListBox1.Items.Add("Can NOT Delete old file: " & xFolder)
                xNoError = False
            End Try

            '---- File write ----
            Try
                Dim fs As New FileStream(xFilePath, FileMode.Create, FileAccess.Write)
                Dim s As New StreamWriter(fs)
                s.WriteLine("Test file, does folder allow write?")
                s.WriteLine("File written:  " & Now.ToString("MM/dd/yyyy hh:mm:ss tt"))
                s.Close()
            Catch ex As Exception
                ListBox1.Items.Add("Can NOT write file: " & xFolder)
                xNoError = False
            End Try

        Else
            ListBox1.Items.Add("Folder NOT found: " & xFolder)
            xNoError = False
        End If

        '==== ELSH ==== ==== ELSH ==== ==== ELSH ==== ==== ELSH ==== ====
        xFolder = BackupBase & "ELSH"
        xFilePath = xFolder & "\Test.txt"
        If Directory.Exists(xFolder) Then
            '==== Folder exist, can we write & delete? ==== ==== ====
            '---- File Delete ----
            Try
                If File.Exists(xFilePath) Then
                    Kill(xFilePath)
                End If
            Catch ex As Exception
                ListBox1.Items.Add("Can NOT Delete old file: " & xFolder)
                xNoError = False
            End Try

            '---- File write ----
            Try
                Dim fs As New FileStream(xFilePath, FileMode.Create, FileAccess.Write)
                Dim s As New StreamWriter(fs)
                s.WriteLine("Test file, does folder allow write?")
                s.WriteLine("File written:  " & Now.ToString("MM/dd/yyyy hh:mm:ss tt"))
                s.Close()
            Catch ex As Exception
                ListBox1.Items.Add("Can NOT write file: " & xFolder)
                xNoError = False
            End Try

        Else
            ListBox1.Items.Add("Folder NOT found: " & xFolder)
            xNoError = False
        End If

        Return xNoError

    End Function   'FolderTest()

    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 
    '==== Command buttons & Other Controls = ==== ==== ==== ==== ==== ==== ====
    '==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== ==== 

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim envString As String
        Dim index As Integer = 1

        envString = Environ(index)
        While envString <> ""

            ListBox1.Items.Add(index & "   " & envString)

            'Debug.WriteLine(index & "   " & envString)

            '---- Next
            index += 1
            envString = Environ(index)
        End While

    End Sub   'Button1_Click

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If Environment_IsDev() Then
            If MachineState = "Finished" Then
                lblAppStatus.Text = "Finished, X-out when you like."
            Else
                '---- Display MachineState & Count ----
                lblAppStatus.Text = MachineState & " in : " & TimerCount
            End If
        Else
            '---- Display MachineState & Count ----
            lblAppStatus.Text = MachineState & " in : " & TimerCount
        End If

        '---- Count Down ----
        If TimerDirection = "Up" Then
            TimerCount += 1
            If TimerCount > TimerLimitHigh Then
                '---- Stop count here, turn off timer ----
                TimerCount = 0
                Timer1.Enabled = False
            End If
        ElseIf TimerDirection = "Down" Then
            TimerCount -= 1
            If TimerCount < TimerLimitLow Then
                '---- Avoid negative count, turn off timer ----
                TimerCount = 0
                Timer1.Enabled = False
  
                '---- What's next? ----
                If MachineState = "Starting" Then
                    Export_ProcessMain()                    'Timer1_Tick()
                ElseIf MachineState = "Processing" Then
                    '---- NO action here ----
                ElseIf MachineState = "Closing" Then
                    '---- Time to close this app ----
                    ListBox1.Items.Add("---- ---- ----")
                    ListBox1.Items.Add("Closing app.")
                    lblAppStatus.Text = "Closed"
                    Application.DoEvents()
                    Me.Close()
                End If
            End If
        End If

    End Sub   'Timer1_Tick


End Class
