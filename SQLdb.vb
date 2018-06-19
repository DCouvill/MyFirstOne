Imports System.Data.SqlClient               'Required to define "SqlConnection"
Public Class SQLdb
    ' Code for SQL Server database: PIPDownload
    '
    '==== Database Connection Strings ==== ====
    '   To SQL Server 2008 at ELMHS
    '**** Change to server migration to Vinyu ****
    ' Fr "Data Source=10.12.20.232;" & _
    ' To "Data Source=OMH-ELS-SQL01;" & _
    Protected Const connstr1 As String = _
        "Data Source=OMH-ELS-SQL01;" & _
        "User ID=webappsusr;" & _
        "Password=mentalh;" & _
        "DataBase=HealthInfo"
    Dim connPIPindex As SqlConnection = New SqlConnection(connstr1)
    'Const DBName = "HealthInfo"

    'Dim conn1 As SqlConnection = New SqlConnection(connstr1)
    Dim ConnUsed As SqlConnection = New SqlConnection

    Dim Conn_Good As Boolean
    Public ReadOnly Property ConnectionGood() As Boolean
        Get
            Return Conn_Good
        End Get
    End Property

    Dim DB_Name As String
    Public ReadOnly Property DatabaseName() As String
        Get
            Return DB_Name
        End Get
    End Property

    '---- Basic Variable List -------------
    Dim sqlstm As String
    Dim sqlstm2 As String
    Dim sqlstm3 As String
    Dim errmsg As String
    Dim msg01 As String
    Dim msg02 As String
    Dim msg03 As String
    Dim Outline As String
    Dim OL As String
    Dim LoadComplete As Boolean
    Dim CurrentUser As String
    Dim KeepLooping As Boolean
    Dim AP As Int16
    Dim rc As Int32
    Dim rcb As Boolean
    Dim rb As Boolean
    Dim rf As Double
    Dim rs As String
    Dim rs1 As String
    Dim rs2 As String
    Dim rs3 As String
    Dim rd As Date
    Dim i As Int16
    Dim x As String
    Dim xx As String
    Dim xxx As String
    Dim str01 As String
    Dim str02 As String
    Dim str03 As String
    Dim str04 As String
    Dim Fname As String
    Dim Fpath As String

    '==== General Properties ==== ====
    Dim new_by As String
    Public Property NewBy() As String
        Get
            Return new_by
        End Get
        Set(ByVal value As String)
            new_by = value
        End Set
    End Property
    '==== MAIN table ==== ====
    Dim pip_UI As String = ""
    Public Property pipUI() As String
        Get
            Return pip_UI
        End Get
        Set(ByVal value As String)
            pip_UI = value
        End Set
    End Property
    Dim pip_UIhex As String
    Public Property pipUIhex() As String
        Get
            Return pip_UIhex
        End Get
        Set(ByVal value As String)
            pip_UIhex = value
        End Set
    End Property

    Dim pip_NEWUID As Int64 = 0
    Public Property pipNEWUID() As Int64
        Get
            Return pip_NEWUID
        End Get
        Set(ByVal value As Int64)
            pip_NEWUID = value
        End Set
    End Property
    Dim pip_SQLUID As Int64 = 0
    Public Property pipSQLUID() As Int64
        Get
            Return pip_SQLUID
        End Get
        Set(ByVal value As Int64)
            pip_SQLUID = value
        End Set
    End Property
    Dim pip_HOSP_NUM As String = ""
    Public Property pipHOSP_NUM() As String
        Get
            Return pip_HOSP_NUM
        End Get
        Set(ByVal value As String)
            pip_HOSP_NUM = value
        End Set
    End Property
    Dim pip_FNAME As String = ""
    Public Property pipFNAME() As String
        Get
            Return pip_FNAME
        End Get
        Set(ByVal value As String)
            pip_FNAME = value
        End Set
    End Property
    Dim pip_LNAME As String = ""
    Public Property pipLNAME() As String
        Get
            Return pip_LNAME
        End Get
        Set(ByVal value As String)
            pip_LNAME = value
        End Set
    End Property
    Dim pip_MID_INIT As String = ""
    Public Property pipMID_INIT() As String
        Get
            Return pip_MID_INIT
        End Get
        Set(ByVal value As String)
            pip_MID_INIT = value
        End Set
    End Property
    Dim pip_STADDRESS As String
    Public Property pipSTADDRESS() As String
        Get
            Return pip_STADDRESS
        End Get
        Set(ByVal value As String)
            pip_STADDRESS = value
        End Set
    End Property
    Dim pip_CITY As String
    Public Property pipCITY() As String
        Get
            Return pip_CITY
        End Get
        Set(ByVal value As String)
            pip_CITY = value
        End Set
    End Property
    Dim pip_STATE As String
    Public Property pipSTATE() As String
        Get
            Return pip_STATE
        End Get
        Set(ByVal value As String)
            pip_STATE = value
        End Set
    End Property

    Dim pip_ZIP As String
    Public Property pipZIP() As String
        Get
            Return pip_ZIP
        End Get
        Set(ByVal value As String)
            pip_ZIP = value
        End Set
    End Property
    Dim pip_RACE As String = ""
    Public Property pipRACE() As String
        Get
            Return pip_RACE
        End Get
        Set(ByVal value As String)
            pip_RACE = value
        End Set
    End Property
    Dim pip_ETHNIC As String = ""
    Public Property pipETHNIC() As String
        Get
            Return pip_ETHNIC
        End Get
        Set(ByVal value As String)
            pip_ETHNIC = value
        End Set
    End Property
    Dim pip_SEX As String = ""
    Public Property pipSEX() As String
        Get
            Return pip_SEX
        End Get
        Set(ByVal value As String)
            pip_SEX = value
        End Set
    End Property
    Dim pip_BIRTHDATE As Date
    Public Property pipBIRTHDATE() As String        'Date
        Get
            Return pip_BIRTHDATE
        End Get
        Set(ByVal value As String)
            pip_BIRTHDATE = value
        End Set
    End Property
    Dim pip_SSN As String = ""
    Public Property pipSSN() As String
        Get
            Return pip_SSN
        End Get
        Set(ByVal value As String)
            pip_SSN = value
        End Set
    End Property
    Dim pip_UNIT As String = ""
    Public Property pipUNIT() As String
        Get
            Return pip_UNIT
        End Get
        Set(ByVal value As String)
            pip_UNIT = value
        End Set
    End Property
    Dim pip_STATUS As String = ""
    Public Property pipSTATUS() As String
        Get
            Return pip_STATUS
        End Get
        Set(ByVal value As String)
            pip_STATUS = value
        End Set
    End Property
    Dim pip_ADM_DATE As Date
    Public Property pipADM_DATE() As String     'Date
        Get
            Return pip_ADM_DATE
        End Get
        Set(ByVal value As String)
            pip_ADM_DATE = value
        End Set
    End Property
    Dim pip_ADM_TIME As String = ""
    Public Property pipADM_TIME() As String
        Get
            Return pip_ADM_TIME
        End Get
        Set(ByVal value As String)
            pip_ADM_TIME = value
        End Set
    End Property
    Dim pip_ADM_TYPE As String
    Public Property pipADM_TYPE() As String
        Get
            Return pip_ADM_TYPE
        End Get
        Set(ByVal value As String)
            pip_ADM_TYPE = value
        End Set
    End Property
    Dim pip_DC_DATE As Date
    Public Property pipDC_DATE() As String  'Date
        Get
            Return pip_DC_DATE
        End Get
        Set(ByVal value As String)
            pip_DC_DATE = value
        End Set
    End Property
    Dim pip_DC_TIME As String = ""
    Public Property pipDC_TIME() As String
        Get
            Return pip_DC_TIME
        End Get
        Set(ByVal value As String)
            pip_DC_TIME = value
        End Set
    End Property
    'LAST_DC  pip_LAST_DC
    Dim pip_LAST_DC As String = ""
    Public Property pipLAST_DC As String
        Get
            Return pip_LAST_DC
        End Get
        Set(ByVal value As String)
            pip_LAST_DC = value
        End Set
    End Property

    Dim pip_FACID As String = ""
    Public Property pipFACID() As String
        Get
            Return pip_FACID
        End Get
        Set(ByVal value As String)
            pip_FACID = value
        End Set
    End Property

    Dim pip_ADMLEGSTAT As String = ""
    Public Property pipADMLEGSTAT As String
        Get
            Return pip_ADMLEGSTAT
        End Get
        Set(ByVal value As String)
            pip_ADMLEGSTAT = value
        End Set
    End Property
    Dim pip_MOD_DATE As Date
    Public Property pipMOD_DATE() As Date
        Get
            Return pip_MOD_DATE
        End Get
        Set(ByVal value As Date)
            pip_MOD_DATE = value
        End Set
    End Property
    Dim pip_MOD_TIME As String = ""
    Public Property pipMOD_TIME() As String
        Get
            Return pip_MOD_TIME
        End Get
        Set(ByVal value As String)
            pip_MOD_TIME = value
        End Set
    End Property

    Dim pip_Div As String = ""
    Public Property pipDIV() As String
        Get
            Return pip_Div
        End Get
        Set(ByVal value As String)
            pip_Div = value
        End Set
    End Property

    'Dim pip_xxx As String = ""
    'Public Property pipxxx() As String
    '    Get
    '        Return pip_xxx
    '    End Get
    '    Set(ByVal value As String)
    '        pip_xxx = value
    '    End Set
    'End Property

    Dim ErrorHandle1 As ErrorHandle = New ErrorHandle   'Save messages to:  DataBase = PIPlog

    '==== New() ==== New() ==== New() ==== ====
    Public Sub New()

        ConnUsed = connPIPindex

        Conn_Good = TestConnection()


    End Sub   'New()

    Public Function TestConnection() As Boolean
        ' Test connection to database

        '---- SQL Statement ---- 
        sqlstm = ""
        sqlstm &= "SELECT "
        sqlstm &= "    * "
        sqlstm &= "FROM "
        sqlstm &= "    tbl_about "

        Dim dsTestConn As DataSet = New DataSet()
        Dim adap1 As SqlDataAdapter = New SqlDataAdapter()
        Dim rcCnt As Int16 = 0

        Try
            dsTestConn.Clear()
            ConnUsed.Open()
            Try
                '==== DataAdapter & DataSets ====
                adap1 = New SqlDataAdapter(sqlstm, ConnUsed)
                adap1.Fill(dsTestConn, "Test")
                rcCnt = dsTestConn.Tables("Test").Rows.Count()
                DB_Name = ConnUsed.Database
                ConnUsed.Close()
                Conn_Good = True
                Return rcCnt
            Catch ex As Exception
                errmsg = "--------------------------------------" & vbCrLf
                errmsg &= ex.ToString & vbCrLf & vbCrLf
                errmsg &= "In SQLdb . TestConnection() A"
#If DEBUG Then
                'MsgBox(errmsg)
                Debug.WriteLine(errmsg)
#End If
                '---- Save Error Message ----
                ErrorHandle1.ClassName = "SQLdb"
                ErrorHandle1.ProcName = "TestConnection"
                ErrorHandle1.ErrorMessage = ex.Message
                ErrorHandle1.ErrorNote = "Part A"
                rc = ErrorHandle1.SaveError()
                '----
                ConnUsed.Close()
                Conn_Good = False
                DB_Name = ""
                Return -1
            End Try
        Catch ex As Exception
            errmsg = "--------------------------------------" & vbCrLf
            errmsg &= ex.ToString & vbCrLf & vbCrLf
            errmsg &= "Connection to database failed." & vbCrLf
            errmsg &= "In SQLdb . TestConnection() B"
#If DEBUG Then
            'MsgBox(errmsg)
            Debug.WriteLine(errmsg)
#End If
            '---- Save Error Message ----
            ErrorHandle1.ClassName = "SQLdb"
            ErrorHandle1.ProcName = "TestConnection"
            ErrorHandle1.ErrorMessage = ex.Message
            ErrorHandle1.ErrorNote = "Part B"
            rc = ErrorHandle1.SaveError()
            ConnUsed.Close()
            DB_Name = ""
            Conn_Good = False
            Return -1
        End Try
    End Function   'TestConnection()
    '==== General ==== ====



    '==== MAIN Table Import ==== ==== MAIN Table Import ==== ====
    ' Fields in : HealthInfo . tbl_PIPindex
    '
    '   Field           Note
    '-------------------------------------------------
    '   pip_id          ID, Automatic
    '   HOSP_NUM
    '   HOSP_NUMkey     Like HOSP_NUM but with hospital divison imbedded in number string
    '                       Example: CHD000112   only for div CHD & SFF
    '   LNAME
    '   FNAME
    '   MID_INIT
    '   RACE
    '   ETHNIC
    '   SEX
    '   BIRTHDATE
    '   SSN
    '   UNIT
    '   STATUS
    '   ADM_DATE
    '   ADM_TYPE
    '   DC_DATE
    '   FACID
    '   Div
    '   Update_DT       Automatic

    Public Function MAIN_SaveOne() As Integer
        ' Save on record to SQL Server table:  "tblmain"

        Dim xHospNumKey As String = ""

        If pip_Div = "SFF" Then
            xHospNumKey = "SFF" & Right(pip_HOSP_NUM, 6)
        ElseIf pip_Div = "CHD" Then
            xHospNumKey = "CHD" & Right(pip_HOSP_NUM, 6)
        Else
            '---- Div most be "FFF" or "ELSH" ----
            xHospNumKey = pip_HOSP_NUM
        End If





        '----
        Dim GivenDate1 As String = ""
        'If pip_BIRTHDATE <= #1/1/1850# Then
        '    GivenDate1 = "1901-01-01 00:00:00"
        'Else
        '    GivenDate1 = "CONVERT(DATETIME, '"
        '    GivenDate1 &= pip_BIRTHDATE.ToString("yyyy-MM-dd")
        '    GivenDate1 &= " 00:00:00', 102)"
        'End If

        If DateValue(pip_BIRTHDATE) <= #1/1/1850# Then
            GivenDate1 = "1/1/1900"
        Else
            GivenDate1 = pip_BIRTHDATE
        End If
        '----
        Dim GivenDate5 As String = ""
        If DateValue(pip_LAST_DC) <= #1/1/1850# Then
            GivenDate5 = "1/1/1900"
        Else
            GivenDate5 = pip_LAST_DC
        End If

        '----
        Dim GivenDate2 As String = ""
        'If pip_ADM_DATE <= #1/1/1900# Then
        '    GivenDate2 = "1901-01-01 00:00:00"
        'Else
        '    GivenDate2 = "CONVERT(DATETIME, '"
        '    GivenDate2 &= pip_ADM_DATE.ToString("yyyy-MM-dd")
        '    GivenDate2 &= " 00:00:00', 102)"
        'End If
        If DateValue(pip_ADM_DATE) < #1/1/1900# Then
            GivenDate2 = "1/1/1900"
        Else
            GivenDate2 = pip_ADM_DATE
        End If



        '----
        Dim GivenDate3 As String = ""
        'If pip_DC_DATE <= #1/1/1900# Then  'Check for date out of range
        '    GivenDate3 = "1901-01-01 00:00:00"
        'Else
        '    GivenDate3 = "CONVERT(DATETIME, '"
        '    GivenDate3 &= pip_DC_DATE.ToString("yyyy-MM-dd")
        '    GivenDate3 &= " 00:00:00', 102)"
        'End If
        If DateValue(pip_DC_DATE) < #1/1/1900# Then
            GivenDate3 = "1/1/1900"
        Else
            GivenDate3 = pip_DC_DATE
        End If


        '----
        Dim GivenDate4 As String = ""
        'If pip_MOD_DATE <= #1/1/1900# Then  'Check for date out of range
        '    GivenDate4 = "1901-01-01 00:00:00"
        'Else
        '    GivenDate4 = "CONVERT(DATETIME, '"
        '    GivenDate4 &= pip_MOD_DATE.ToString("yyyy-MM-dd")
        '    GivenDate4 &= " 00:00:00', 102)"
        'End If
        If DateValue(pip_MOD_DATE) < #1/1/1900# Then
            GivenDate4 = "1/1/1900"
        Else
            GivenDate4 = pip_MOD_DATE
        End If


        '---- Quote Mark fix ----
        pip_LNAME = Replace(pip_LNAME, "'", "''")
        pip_LNAME = Trim(pip_LNAME)

        pip_FNAME = Replace(pip_FNAME, "'", "''")
        pip_FNAME = Trim(pip_FNAME)

        pip_UI = Replace(pip_UI, "'", "''")
        'pip_UI = Trim(pip_UI)

        pip_STADDRESS = Replace(pip_STADDRESS, "'", "''")
        pip_STADDRESS = Trim(pip_STADDRESS)

        pip_CITY = Replace(pip_CITY, "'", "''")
        pip_CITY = Trim(pip_CITY)

        pip_ZIP = Trim(pip_ZIP)

        '---- Diagonstic, trap suspected Hospital Number ----
        'If pip_HOSP_NUM = "000056819" Then
        '    i = 0
        'End If
        '---- SQL Statement ---- ----
        sqlstm = ""
        sqlstm &= "INSERT INTO "
        sqlstm &= "[tbl_PIPindex] "                    'Production
        'sqlstm &= "[tbl_PIPindex_2016_01_27empty] "     'Test
        sqlstm &= "("
        'sqlstm &= "    [UI], "
        'sqlstm &= "    [NEWUID], "
        'sqlstm &= "    [SQLUID], "
        'sqlstm &= "    [UIhex], "
        sqlstm &= "    [HOSP_NUM], "
        'sqlstm &= "    [HOSP_UNIT], "
        sqlstm &= "    [FNAME], "
        sqlstm &= "    [LNAME], "
        sqlstm &= "    [MID_INIT], "
        'sqlstm &= "    STADDRESS, "
        'sqlstm &= "    CITY, "      '
        'sqlstm &= "    STATE, "     '
        'sqlstm &= "    ZIP, "       '
        sqlstm &= "    [RACE], "
        sqlstm &= "    [ETHNIC], "
        sqlstm &= "    [SEX], "
        sqlstm &= "    [BIRTHDATE], "
        sqlstm &= "    [SSN], "
        sqlstm &= "    [UNIT], "
        sqlstm &= "    [STATUS], "
        sqlstm &= "    [ADM_DATE], "
        'sqlstm &= "    [ADM_TIME], "
        sqlstm &= "    [ADM_TYPE], "
        sqlstm &= "    [DC_DATE], "
        'sqlstm &= "    [DC_TIME], "
        sqlstm &= "    [FACID], "
        'sqlstm &= "    [ADMLEGSTAT], "

        sqlstm &= "    [LAST_DC], "

        'sqlstm &= "    [MOD_DATE], "
        'sqlstm &= "    [MOD_TIME] "
        'sqlstm &= "    [source] "
        sqlstm &= "    [Div], "

        sqlstm &= "    [HOSP_NUMkey] "

        sqlstm &= ") "
        sqlstm &= "VALUES "
        sqlstm &= "("
        'sqlstm &= " '" & sss & "', "   
        'sqlstm &= " '" & pip_UI & "', "
        'sqlstm &= " " & pip_NEWUID & ", "
        'sqlstm &= " " & pip_SQLUID & ", "
        'sqlstm &= " '" & pip_UIhex & "', "
        sqlstm &= " '" & pip_HOSP_NUM & "', "
        'sqlstm &= " '" & pip_HOSP_UNIT & "', "
        sqlstm &= " '" & pip_FNAME & "', "
        sqlstm &= " '" & pip_LNAME & "', "
        sqlstm &= " '" & pip_MID_INIT & "', "
        'sqlstm &= " '" & pip_STADDRESS & "', "
        'sqlstm &= " '" & pip_CITY & "', "
        'sqlstm &= " '" & pip_STATE & "', "
        ' sqlstm &= " '" & pip_ZIP & "', "
        sqlstm &= " '" & pip_RACE & "', "
        sqlstm &= " '" & pip_ETHNIC & "', "
        sqlstm &= " '" & pip_SEX & "', "
        sqlstm &= " '" & GivenDate1 & "', "       'pip_BIRTHDATE
        sqlstm &= " '" & pip_SSN & "', "
        sqlstm &= " '" & pip_UNIT & "', "
        sqlstm &= " '" & pip_STATUS & "', "
        sqlstm &= " '" & GivenDate2 & "', "       'pip_ADM_DATE
        'sqlstm &= " '" & pip_ADM_TIME & "', "
        sqlstm &= " '" & pip_ADM_TYPE & "', "
        sqlstm &= " '" & GivenDate3 & "', "       'pip_DC_DATE
        'sqlstm &= " '" & pip_DC_TIME & "', "
        sqlstm &= " '" & pip_FACID & "', "
        'sqlstm &= " '" & pip_ADMLEGSTAT & "', "
        sqlstm &= " '" & GivenDate5 & "', "       'pip_LAST_DC



        'sqlstm &= " " & GivenDate4 & ", "       'pip_MOD_DATE
        'sqlstm &= " '" & pip_MOD_TIME & "' "
        'sqlstm &= " '" & pip_source & "' "  pip_Div

        sqlstm &= " '" & pip_Div & "', "

        sqlstm &= " '" & xHospNumKey & "' "

        sqlstm &= ")"
        sqlstm &= "    "

        'Debug.WriteLine(sqlstm)
        'ConnUsed = conn1
        Dim rCnt As Int16
        Dim myCmd As New SqlCommand()
        ConnUsed.Open()
        Try
            '---- Create & Run the command
            myCmd.CommandText = sqlstm
            myCmd.CommandType = CommandType.Text
            myCmd.Connection = ConnUsed
            rCnt = myCmd.ExecuteNonQuery()
            ConnUsed.Close()
            Return rCnt
        Catch ex As Exception
            '---- Display error
            errmsg = "-----------------------------------------------" & vbCrLf
            errmsg &= ex.Message & vbCrLf & vbCrLf
            errmsg &= pip_FNAME & " " & pip_LNAME & vbCrLf
            errmsg &= "Hospital Number:  " & pip_HOSP_NUM & vbCrLf
            errmsg &= "SQLdb . MAIN_SaveOne() " & vbCrLf
#If DEBUG Then
            'MsgBox(errmsg)
            Debug.WriteLine(sqlstm)
            Debug.WriteLine(errmsg)
#End If
            '---- Save Error Message ----
            ErrorHandle1.ClassName = "SQLdb"
            ErrorHandle1.ProcName = "MAIN_SaveOne"
            ErrorHandle1.ErrorMessage = ex.Message
            ErrorHandle1.ErrorNote = ConnUsed.Database
            rc = ErrorHandle1.SaveError()
            '----
            ConnUsed.Close()
            '---- Return failure value, False
            Return -1
        End Try
    End Function   'MAIN_SaveOne()

    Public Function PIPindex_Clear() As Integer
        ' Clear / Delete all records in table:  tbl_PIPindex

        'For DBCC to work you have to set permission for the user "WebAppUser"
        'In SSMS >> Security >> logins
        '       Find the user in Logins
        '>> User Mapping
        '       Find the database name
        'check "db_owner"

        '---- SQL Statement ----tbl_PIPindex_2016_01_27
        sqlstm = ""
        sqlstm &= "DELETE "
        sqlstm &= "FROM "
        sqlstm &= "[tbl_PIPindex] "
        sqlstm &= "DBCC CHECKIDENT(tbl_PIPindex, RESEED, 0)"

        'sqlstm &= "DELETE "
        'sqlstm &= "FROM "
        'sqlstm &= "[tbl_PIPindex_2016_01_27] "
        'sqlstm &= "DBCC CHECKIDENT(tbl_PIPindex_2016_01_27, RESEED, 0)"

        sqlstm &= ""

        Dim cmdDelRec As SqlCommand
        Dim resultCnt As Integer
        Try
            ConnUsed.Open()
            cmdDelRec = New SqlCommand(sqlstm)
            cmdDelRec.Connection = ConnUsed
            resultCnt = cmdDelRec.ExecuteNonQuery()
            ConnUsed.Close()
            Return resultCnt
        Catch ex As Exception
            errmsg = "--------------------------------------" & vbCrLf
            errmsg &= ex.Message & vbCrLf
            errmsg &= errmsg & "Trying to clear MAIN." & vbCrLf
            errmsg &= errmsg & "MAIN_Clear()"
#If DEBUG Then
            'MsgBox(errmsg)
            Debug.WriteLine(errmsg)
#End If
            '---- Save Error Message ----
            ErrorHandle1.ClassName = "SQLdb"
            ErrorHandle1.ProcName = "PIPindex_Clear"
            ErrorHandle1.ErrorMessage = ex.Message
            ErrorHandle1.ErrorNote = ConnUsed.Database
            rc = ErrorHandle1.SaveError()
            '----
            resultCnt = -1
            ConnUsed.Close()
            Return resultCnt
        End Try
    End Function 'PIPindex_Clear()

    Public Function MAIN_Clear() As Integer
        ' Clear all records for that database from table: "MAIN"

        '---- SQL Statement ----
        sqlstm = ""
        sqlstm &= "DELETE "
        sqlstm &= "FROM "
        sqlstm &= "    MAIN "
        'sqlstm &= "    "

        Dim cmdDelRec As SqlCommand
        Dim resultCnt As Integer
        Try
            ConnUsed.Open()
            cmdDelRec = New SqlCommand(sqlstm)
            cmdDelRec.Connection = ConnUsed
            resultCnt = cmdDelRec.ExecuteNonQuery()
            ConnUsed.Close()
            Return resultCnt
        Catch ex As Exception
            errmsg = "--------------------------------------" & vbCrLf
            errmsg &= ex.Message & vbCrLf
            errmsg &= errmsg & "Trying to clear MAIN." & vbCrLf
            errmsg &= errmsg & "MAIN_Clear()"
#If DEBUG Then
            'MsgBox(errmsg)
            Debug.WriteLine(errmsg)
#End If
            '---- Save Error Message ----
            ErrorHandle1.ClassName = "SQLdb"
            ErrorHandle1.ProcName = "MAIN_Clear"
            ErrorHandle1.ErrorMessage = ex.Message
            ErrorHandle1.ErrorNote = ConnUsed.Database
            rc = ErrorHandle1.SaveError()
            '----
            resultCnt = -1
            ConnUsed.Close()
            Return resultCnt
        End Try
    End Function   'MAIN_Clear()

 


    '==== General purpose functions ==== ====
    Dim WkgStr As String = ""
    Public Function StringCondition(ByVal gString As String) As String
        ' Given a string remove/replace certain characters
        ' Before saving to the database
        Dim WkgCnt As Integer = 10
        Dim SpCnt As Integer = 0

        '---- Trim the string ----
        WkgStr = Trim(gString)
        '---- Quote Fix ----
        WkgStr = Replace(WkgStr, "'", "''")
        WkgStr = Replace(WkgStr, vbCrLf, " ")

        WkgCnt = 5
        SpCnt = Stringcondition_DoubleSpaces(WkgStr)
        Do While SpCnt > 0 And WkgCnt > 0
            SpCnt = Stringcondition_DoubleSpaces(WkgStr)
            WkgCnt -= 1
        Loop
        WkgStr = Trim(WkgStr)

        Return WkgStr

    End Function   'StringCondition()

    Private Function StringCondition_DoubleSpaces(ByVal gString As String) As Integer
        ' Given a string remove Double Spaces
        ' Return 0 if no double spaces remaining
        ' Leave string in "WkgStr"

        Dim SpCnt As Integer = 0
        '---- 1st replace double spaces with a single space ----
        WkgStr = Replace(gString, "  ", " ")

        SpCnt = InStr(WkgStr, "  ")
        Return SpCnt

    End Function   'StringCondition_DoubleSpaces()

End Class
