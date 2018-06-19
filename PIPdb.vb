Imports System.Data.OleDb
Imports System.IO
'Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Diagnostics

Public Class PIPdb
    ' PIP database object
    ' Use to connect to any of the PIP databases
    ' See New() constructor
    '
    '********************************************************************************
    '**** Important Note:                                       *********************
    '**** Requires installation of Support for -- Visual FoxPro *********************
    '**** Download:  Microsoft OLE Provider for Visual FoxPro 9.0 (or later *********
    '**** from:    http://www.microsoft.com/en-us/download/details.aspx?id=14839 ****
    '**** Instruction: double-clidk on downloaded file: "VFPOLEDBSetup.msi" *********
    '**** May need to install on each machine that runs this app ********************
    '**** Copy at C:\SoftwareDons\MS VFPOLEDB\ **************************************
    '****
    '**** Don't forget to install this on the app server         ********************
    '********************************************************************************

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
    Dim AP2 As Int16
    Dim rc As Int32
    Dim rb As Boolean
    Dim rf As Double
    Dim rs As String
    Dim rs1 As String
    Dim rs2 As String
    Dim rs3 As String
    Dim rs4 As String
    Dim i As Int16
    Dim j As Int16
    Dim k As Int16
    Dim x As String
    Dim xx As String
    Dim xxx As String
    Dim ConnUsed As OleDbConnection

    '==== Connections ====
    '---- FFF Production ----
    Protected Const connstr1 As String = _
    "Provider=vfpoledb.1;" & _
    "Data Source=G:\PIP\FFF\up2BR\;"
    Dim connFFF As New OleDbConnection(connstr1)
    '---- FFF Test ----
    Protected Const connstr1test As String = _
    "Provider=vfpoledb.1;" & _
    "Data Source=C:\Data\PIP\FFF\;"
    Dim connFFFtest As New OleDbConnection(connstr1test)

    '---- ELSH Production ----
    Protected Const connstr2 As String = _
    "Provider=vfpoledb.1;" & _
    "Data Source=G:\PIP\ELSH\up2BR\;"
    Dim connELSH As New OleDbConnection(connstr2)
    '---- ELSH Test ----
    Protected Const connstr2test As String = _
       "Provider=vfpoledb.1;" & _
       "Data Source=C:\Data\PIP\ELSH\;"
    Dim connELSHtest As New OleDbConnection(connstr2test)

    '---- Group Homes (GH) Production ----
    Protected Const connstr3 As String = _
        "Provider=vfpoledb.1;" & _
        "Data Source=G:\PIP\GRPHOME\up2br\;"
    Dim connGH As New OleDbConnection(connstr3)
    '---- Group Homes (GH) Test ----
    Protected Const connstr3test As String = _
        "Provider=vfpoledb.1;" & _
        "Data Source=C:\Data\PIP\GRPHOME\;"
    Dim connGHtest As New OleDbConnection(connstr3test)

    '---- Secure Forensic facility (SFF) Production ----
    Protected Const connstr4 As String = _
    "Provider=vfpoledb.1;" & _
    "Data Source=G:\PIP\SFF\up2br\;"
    Dim connSFF As New OleDbConnection(connstr3)
    '---- Group Homes (GH) Test ----
    Protected Const connstr4test As String = _
        "Provider=vfpoledb.1;" & _
        "Data Source=C:\Data\PIP\SFF\;"
    Dim connSFFtest As New OleDbConnection(connstr3test)

    '==== Database ====
    Dim Conn_To As String = ""                      '"FFF" or "ELSH", etc
    Dim MAINcnt As Int64 = 0                        'Records found by DBTest
    Public ReadOnly Property MainRows() As Int64
        Get
            Return MAINcnt
        End Get
    End Property

    Public ReadOnly Property SourceFolder() As String
        Get
            Return ConnUsed.DataSource
        End Get
    End Property

    Dim In_Test As Boolean = False
    Public ReadOnly Property InTest() As Boolean
        Get
            Return In_Test
        End Get
    End Property

    'Dim Inter_Active As Boolean = True      
    '---- Interactive ----
    '   Means the app will prompt user with error messages etc.
    '   You may not want this if running unattended on a server
    '       Windows app (set to True) 
    '       Unattended (set to False)
    Public Property InterActive As Boolean

    Dim ErrorHandle2 As ErrorHandle

    '==== New() ==== New() ==== New() ==== ====
    Public Sub New(ByVal ConnTo As String, ByVal TestDB As Boolean)
        ' Allowed ConnTo:  "ELSH","FFF", or "GH"
        ' Testdb = True will connect with a Test database

        If Not InterActive Then
            ErrorHandle2 = New ErrorHandle
        End If

        In_Test = TestDB

        If ConnTo = "ELSH" Then
            If In_Test Then
                ConnUsed = connELSHtest
            Else
                ConnUsed = connELSH
            End If
            Conn_To = ConnTo
        ElseIf ConnTo = "FFF" Then
            If In_Test Then
                ConnUsed = connFFFtest
            Else
                ConnUsed = connFFF
            End If
            Conn_To = ConnTo
        ElseIf ConnTo = "GH" Then       'Group Homes
            If In_Test Then
                ConnUsed = connGHtest
            Else
                ConnUsed = connGH
            End If
            Conn_To = ConnTo
        ElseIf ConnTo = "SFF" Then
            If In_Test Then
                ConnUsed = connSFFtest
            Else
                ConnUsed = connSFF
            End If
            Conn_To = ConnTo
        Else
            errmsg = "Unknown database " & ConnTo
            ConnUsed = Nothing
            Conn_To = ""
        End If

    End Sub   'New()

    Public Function DBConnTest() As Int16
        ' Test connection to the database
        ' Return number of records found in the UNIT table
        ' Return -1 on error
        '---- SQL Statement ----
        sqlstm = ""
        sqlstm &= "SELECT "
        sqlstm &= "    * "
        sqlstm &= "FROM "
        sqlstm &= "    UNIT "
        'sqlstm &= "    MAIN "      'Takes too long
        'sqlstm &= " "

        Dim dsTest As DataSet
        Dim RecCnt As Int64 = 0
        dsTest = New DataSet()
        Dim adap1 As New OleDbDataAdapter()
        Try
            ConnUsed.Open()
            Try
                '==== DataAdapter & DataSets ====
                adap1 = New OleDbDataAdapter(sqlstm, ConnUsed)
                adap1.Fill(dsTest, "Test")
                RecCnt = dsTest.Tables("Test").Rows.Count()
                ConnUsed.Close()
                dsTest = Nothing
                Return RecCnt
            Catch ex As Exception
                errmsg = "------------------------------------------" & vbCrLf
                errmsg &= ex.ToString & vbCrLf & vbCrLf
                errmsg &= "In PIPdb.DBConnTest()     DB:  " & Conn_To
#If DEBUG Then
                Debug.WriteLine(errmsg)
#End If
                If InterActive Then
                    MsgBox(errmsg)
                Else
                    '---- Save Error Message ----
                    ErrorHandle2.ClassName = "PIPdb"
                    ErrorHandle2.ProcName = "DBConnTest"
                    ErrorHandle2.ErrorMessage = ex.Message
                    ErrorHandle2.ErrorNote = ""
                    rc = ErrorHandle2.SaveError()
                End If
              
                '----
                ConnUsed.Close()
                dsTest = Nothing
                RecCnt = 0
                Return -1
            End Try
        Catch ex As Exception
            errmsg = "------------------------------------------" & vbCrLf
            errmsg &= ex.ToString & vbCrLf & vbCrLf
            errmsg &= "In PIPdb.DBConnTest()     DB:  " & Conn_To
#If DEBUG Then
            Debug.WriteLine(errmsg)
#End If
            If InterActive Then
                MsgBox(errmsg)
            Else
                '---- Save Error Message ----
                ErrorHandle2.ClassName = "PIPdb"
                ErrorHandle2.ProcName = "DBConnTest"
                ErrorHandle2.ErrorMessage = ex.Message
                ErrorHandle2.ErrorNote = ""
                rc = ErrorHandle2.SaveError()
            End If
            '----
            dsTest = Nothing
            MAINcnt = 0
            Return -2
        End Try
    End Function   'DBConnTest() 


    '==== MAIN Table ==== ==== ==== MAIN Table ==== ==== ====
    '                      Target SQL database : HealthInfo
    ' MAIN fields to target SQL database table : tbl_PIPindex
    '   pip_id
    '   HOSP_NUM
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


    'Public dsMAINexist As DataSet = New DataSet
    Public Function MAIN_Exist() As Int64
        ' Does the MAIN table exist & does it have records
        ' return -1 if error or does NOT exist
        ' return file length if found 

        Dim FilePath As String
        Dim FileSize As Long

        FilePath = ConnUsed.DataSource        'Like: "G:\PIP\FFF\up2BR\"
        FilePath &= "MAIN.DBF"
        If File.Exists(FilePath) Then
            Dim MyFile As New FileInfo(FilePath)
            FileSize = MyFile.Length
            If FileSize > 1100 Then
                Return FileSize
            Else
                Return 0
            End If
        Else
            Return -1
        End If
    End Function   'MAIN_Exist()

    Public dsMAINtable As DataSet = New DataSet
    Public Function MAIN_Export() As Integer
        ' Query PIP for ALL MAIN table records
        ' For use in Import to SQL Server

        '---- SQL Statement ----
        sqlstm = ""
        sqlstm &= "SELECT "
        sqlstm &= "    UI, "        '
        sqlstm &= "    NEWUID, "    '
        sqlstm &= "    SQLUID, "    '
        sqlstm &= "    HOSP_NUM, "  '
        sqlstm &= "    FNAME, "     '
        sqlstm &= "    LNAME, "     '
        sqlstm &= "    MID_INIT, "  '
        sqlstm &= "    STADDRESS, " '
        sqlstm &= "    CITY, "      '
        sqlstm &= "    STATE, "     '
        sqlstm &= "    ZIP, "       '
        sqlstm &= "    RACE, "      '
        sqlstm &= "    ETHNIC, "    '
        sqlstm &= "    SEX, "       '
        sqlstm &= "    BIRTHDATE, " '
        sqlstm &= "    SSN, "       '
        sqlstm &= "    UNIT, "      '
        sqlstm &= "    STATUS, "    '
        sqlstm &= "    ADM_DATE, "  '
        sqlstm &= "    ADM_TIME, "  '
        sqlstm &= "    ADM_TYPE, "  '
        sqlstm &= "    DC_DATE, "   '
        sqlstm &= "    DC_TIME, "   '
        sqlstm &= "    FACID, "     '            '00337=FFF, 00332=ELSH
        sqlstm &= "    MOD_DATE, "  '
        sqlstm &= "    MOD_TIME, "   '
        sqlstm &= "    LAST_DC "   '
        sqlstm &= " "
        sqlstm &= "FROM "
        sqlstm &= "    MAIN "
        sqlstm &= ""

        Dim adap1 As New OleDbDataAdapter()
        Dim rcCnt As Int16 = 0
        dsMAINtable.Clear()

        Try
            ConnUsed.Open()
            '==== DataAdapter & DataSets ====
            adap1 = New OleDbDataAdapter(sqlstm, ConnUsed)
            adap1.Fill(dsMAINtable, "MAINtable")
            rcCnt = dsMAINtable.Tables("MAINtable").Rows.Count()
            ConnUsed.Close()
            Return rcCnt
        Catch ex As Exception
            errmsg = "--------------------------------------" & vbCrLf
            errmsg &= Err.Number & vbCrLf
            errmsg &= ex.Source & vbCrLf
            errmsg &= ex.ToString & vbCrLf & vbCrLf
            errmsg &= "In PIPdb.MAIN_Import()     DB:  " & Conn_To
            ConnUsed.Close()
#If DEBUG Then
            Debug.WriteLine(errmsg)
#End If
            If InterActive Then
                MsgBox(errmsg)
            Else
                '---- Save Error Message ----
                ErrorHandle2.ClassName = "PIPdb"
                ErrorHandle2.ProcName = "MAIN_Import"
                ErrorHandle2.ErrorMessage = ex.Message
                ErrorHandle2.ErrorNote = ""
                rc = ErrorHandle2.SaveError()
            End If
            '----
            rcCnt = 0
            Return -1
        End Try
    End Function  'MAIN_Import()

    Public dsTableDS As DataSet = New DataSet
    Public Function ExportTable(ByVal gTable As String) As Integer
        ' Generic function to export any PIP table
        ' gTable is the complete file path to the table like: "G:\PIP\FFF\up2br\MAIN.DBF"

        '---- SQL Statement ----
        sqlstm = ""
        sqlstm &= "SELECT "
        sqlstm &= "    * "        '
        sqlstm &= "FROM " & gTable
        sqlstm &= ""

        Dim adap1 As New OleDbDataAdapter()
        Dim rcCnt As Int64 = 0
        dsTableDS.Clear()

        Try
            ConnUsed.Open()     'Error:  The 'vfpoledb.1' provider is not registered on the local machine.
            'Error:  Invalid path or file name.
            '==== DataAdapter & DataSets ====
            adap1 = New OleDbDataAdapter(sqlstm, ConnUsed)
            adap1.Fill(dsTableDS, "xTable")
            rcCnt = dsTableDS.Tables("xTable").Rows.Count()
            ConnUsed.Close()
            Return rcCnt
        Catch ex As Exception
            Debug.WriteLine(sqlstm)
            errmsg = "--------------------------------------" & vbCrLf
            errmsg &= Err.Number & vbCrLf
            errmsg &= ex.Source & vbCrLf
            errmsg &= ex.ToString & vbCrLf & vbCrLf
            errmsg &= "In PIPdb.ExportTable()     DB:  " & Conn_To
            ConnUsed.Close()
#If DEBUG Then
            Debug.WriteLine(errmsg)
#End If
            If InterActive Then
                MsgBox(errmsg)
            Else
                '---- Save Error Message ----
                ErrorHandle2.ClassName = "PIPdb"
                ErrorHandle2.ProcName = "ExportTable"
                ErrorHandle2.ErrorMessage = ex.Message
                ErrorHandle2.ErrorNote = ""
                rc = ErrorHandle2.SaveError()
            End If
            '----
            rcCnt = 0
            Return -1
        End Try
    End Function   'ExportTable()
 
 

   
End Class
