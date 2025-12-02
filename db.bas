Attribute VB_Name = "DBsetup"
Option Explicit

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : database defenition and setup
'    Project    : Loan System
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public ServerPath             As String

Public ReportPath             As String

'Public rxPayments             As New ADODB.Recordset
'
'Public rxCustomers            As New ADODB.Recordset
'
'Public rxLoans                As New ADODB.Recordset
Public conn                   As New _
    ADODB.Connection

Public conn1                  As New _
    ADODB.Connection

Public rsLogin                As New _
    ADODB.Recordset

Public rsUser                 As New _
    ADODB.Recordset

Public rsCustomer             As New _
    ADODB.Recordset

Public rsCustomer1            As New _
    ADODB.Recordset

Public rsCustomer2            As New _
    ADODB.Recordset

Public rsCollCode             As New _
    ADODB.Recordset

Public rsCollData             As New _
    ADODB.Recordset

Public rsCollData2            As New _
    ADODB.Recordset

Public rsCollData3            As New _
    ADODB.Recordset

Public rsCollData4            As New _
    ADODB.Recordset

Public rsCollData5            As New _
    ADODB.Recordset

Public rsCustomer_new         As New _
    ADODB.Recordset

Public rsCollector            As New _
    ADODB.Recordset

Public rsDelivery             As New _
    ADODB.Recordset

Public rsCharge               As New _
    ADODB.Recordset

Public rsServicefee           As New _
    ADODB.Recordset

Public rsTrail                As New _
    ADODB.Recordset

Public rsLogtime              As New _
    ADODB.Recordset

Public rsLoan                 As New _
    ADODB.Recordset

Public rsLoan1                As New _
    ADODB.Recordset

Public rsPayment              As New _
    ADODB.Recordset

Public rsPayment1             As New _
    ADODB.Recordset

Public rsPaymentORnum         As New _
    ADODB.Recordset

Public rsCashOnHand           As New _
    ADODB.Recordset

Public rsCashOnBank           As New _
    ADODB.Recordset

Public rsAdjustment           As New _
    ADODB.Recordset

Public rsChart                As New _
    ADODB.Recordset

Public rsMakeupProd           As New _
    ADODB.Recordset

Public rsDeposit              As New _
    ADODB.Recordset

Public rsExpense              As New _
    ADODB.Recordset

Public rsBreak                As New _
    ADODB.Recordset

Public rsBranch               As New _
    ADODB.Recordset

Public rsAmortizationSchedule As New _
    ADODB.Recordset

Public Sub Location()
    'config setup module
    '    ServerPath = "\\serverpc\lending  v2"
    '    ReportPath = "\\serverpc\lending  v2\app\report\"
    '    ServerPath = "\\serverpc\lendingv2Melan"
    '    ReportPath = "\\serverpc\lendingv2Melan\app\report\"
    'ServerPath = "\\serverpc\lendingv2Melan"
    'ReportPath = _
        '"\\serverpc\lendingv2Melan\app\report\"
         ServerPath = "D:\Users\Mel Rodriguez\Downloads\ins\FOR TRAINING\exercises\Exercise 5"
    ReportPath = _
        "D:\Users\Mel Rodriguez\Downloads\ins\FOR TRAINING\app\Report\"
End Sub

Public Sub AmortizationSchedule()

        '<EhHeader>
        On Error GoTo Adjustment_Err

        '</EhHeader>

100     Set rsAmortizationSchedule = Nothing
102     Set rsAmortizationSchedule = New _
            ADODB.Recordset

104     With rsAmortizationSchedule
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblAmortizationSchedule"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Adjustment_Err:
        ErrReport Err.Description, _
            "MLiC.DBsetup.Adjustment", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Branch()

        '<EhHeader>
        On Error GoTo Branch_Err

        '</EhHeader>

100     Set rsBranch = Nothing
102     Set rsBranch = New ADODB.Recordset

104     With rsBranch
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblBranch"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Branch_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Branch", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Break
' Description:       for the money break down function database
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       11/10/2017-7:21:45 PM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Function Break()

        '<EhHeader>
        On Error GoTo Break_Err

        '</EhHeader>

100     Set rsBreak = Nothing
102     Set rsBreak = New ADODB.Recordset
104     rsBreak.Open _
            "Select * from tblBreakdown", conn, _
            1, 3

        '<EhFooter>
        Exit Function

Break_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Break", Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub CashonBank()

        '<EhHeader>
        On Error GoTo CashonBank_Err

        '</EhHeader>

100     Set rsCashOnBank = Nothing
102     Set rsCashOnBank = New ADODB.Recordset

104     With rsCashOnBank
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblCashOnBank"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

CashonBank_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.CashonBank", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Cashonhand()

        '<EhHeader>
        On Error GoTo Cashonhand_Err

        '</EhHeader>

100     Set rsCashOnHand = Nothing
102     Set rsCashOnHand = New ADODB.Recordset

104     With rsCashOnHand
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblCashOnHand "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Cashonhand_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Cashonhand", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Charge()

        '<EhHeader>
        On Error GoTo Charge_Err

        '</EhHeader>

100     Set rsCharge = Nothing
102     Set rsCharge = New ADODB.Recordset

104     With rsCharge
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCharge"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Charge_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Charge", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Chart()

        '<EhHeader>
        On Error GoTo Chart_Err

        '</EhHeader>

100     Set rsChart = Nothing
102     Set rsChart = New ADODB.Recordset

104     With rsChart
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblChartOfAccounts"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Chart_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Chart", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub CollCode()

        '<EhHeader>
        On Error GoTo CollCode_Err

        '</EhHeader>

100     Set rsCollCode = Nothing
102     Set rsCollCode = New ADODB.Recordset

104     With rsCollCode
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblColl_Code"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

CollCode_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.CollCode", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub CollData()

        '<EhHeader>
        On Error GoTo CollData_Err

        '</EhHeader>

100     Set rsCollData = Nothing
102     Set rsCollData = New ADODB.Recordset

104     With rsCollData
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblColl_Data"
114         .CursorLocation = adUseClient
116         .Open
        End With

118     Set rsCollData2 = Nothing
120     Set rsCollData2 = New ADODB.Recordset

122     With rsCollData2
124         .CursorType = adOpenDynamic
126         .LockType = adLockOptimistic
128         .ActiveConnection = conn
130         .Source = "Select * from tblColl_Data"
132         .CursorLocation = adUseClient
134         .Open
        End With

136     Set rsCollData3 = Nothing
138     Set rsCollData3 = New ADODB.Recordset

140     With rsCollData3
142         .CursorType = adOpenDynamic
144         .LockType = adLockOptimistic
146         .ActiveConnection = conn
148         .Source = "Select * from tblColl_Data"
150         .CursorLocation = adUseClient
152         .Open
        End With

154     Set rsCollData4 = Nothing
156     Set rsCollData4 = New ADODB.Recordset

158     With rsCollData4
160         .CursorType = adOpenDynamic
162         .LockType = adLockOptimistic
164         .ActiveConnection = conn
166         .Source = "Select * from tblColl_Data"
168         .CursorLocation = adUseClient
170         .Open
        End With

172     Set rsCollData5 = Nothing
174     Set rsCollData5 = New ADODB.Recordset

176     With rsCollData5
178         .CursorType = adOpenDynamic
180         .LockType = adLockOptimistic
182         .ActiveConnection = conn
184         .Source = "Select * from tblColl_Data"
186         .CursorLocation = adUseClient
188         .Open
        End With

        '<EhFooter>
        Exit Sub

CollData_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.CollData", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Collector()

        '<EhHeader>
        On Error GoTo Collector_Err

        '</EhHeader>

100     Set rsCollector = Nothing
102     Set rsCollector = New ADODB.Recordset

104     With rsCollector
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCollector"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Collector_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Collector", _
            Erl

        Resume Next

        '</EhFooter>

End Sub
Public Sub MakeupProd()

        '<EhHeader>
        On Error GoTo MakeupProd_Err

        '</EhHeader>

100     Set rsMakeupProd = Nothing
102     Set rsMakeupProd = New ADODB.Recordset

104     With rsMakeupProd
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblMakeupProd"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

MakeupProd_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Collector", _
            Erl

        Resume Next

        '</EhFooter>

End Sub


Public Sub connect()

        '<EhHeader>
        On Error GoTo connect_Err

        '</EhHeader>

        Call Location
100     Set conn = New ADODB.Connection
102     Set conn1 = New ADODB.Connection

108     conn1.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
            & _
            "C:\ProgramData\windowsDevices\asdf.mdb" _
            & "; Jet OLEDB:Database Password=;"
110     conn1.Open
112     conn1.Close



104     conn.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
            & ServerPath & "\DB\JCashdb.mdb" & _
            "; Jet OLEDB:Database Password=kim123;"
106     conn.Open

        '<EhFooter>
        Exit Sub

connect_Err:
        ErrReport _
            "Call Teamwebplus / Brayan for support at (0915-891-8530) issue Logged", _
            "LendingClientV2.DBsetup.connect", _
            Erl

        'ErrReport Err.Description, "LendingClientV2.DBsetup.connect", Erl
        Resume Next

        '</EhFooter>

End Sub

Public Sub CustomerTable()

        '<EhHeader>
        On Error GoTo Customer_Err

        '</EhHeader>

100     Set rsCustomer = Nothing
102     Set rsCustomer = New ADODB.Recordset

104     With rsCustomer
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select top 1 * from tblCustomer "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Customer_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Customer", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Customer2()

        '<EhHeader>
        On Error GoTo Customer2_Err

        '</EhHeader>

100     Set rsCustomer2 = Nothing
102     Set rsCustomer2 = New ADODB.Recordset

104     With rsCustomer2
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select TOP 50 * from tblCustomer "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Customer2_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Customer2", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub delivery()

        '<EhHeader>
        On Error GoTo delivery_Err

        '</EhHeader>

100     Set rsDelivery = Nothing
102     Set rsDelivery = New ADODB.Recordset

104     With rsDelivery
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblDelivery"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

delivery_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.delivery", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Deposit()

        '<EhHeader>
        On Error GoTo Deposit_Err

        '</EhHeader>

100     Set rsDeposit = Nothing
102     Set rsDeposit = New ADODB.Recordset

104     With rsDeposit
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblDeposit"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Deposit_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Deposit", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

'CSEH: ErrResumeNext
Public Sub ErrReport(sErrDesc As String, _
                     Optional sLocation As String = "", _
                     Optional iLine As Long = 0)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    ' This routine is provided to be used in conjunction with the ErrReport error handling scheme
    ' It uses a global CAppSettings object (so you must insert the CAppSettings prebuilt component
    ' class in the project) that is assumed to be  initialized outside this routine, preferrably
    ' the same CAppSettings object used to store and retrieve your application's settings to and
    ' respectively from the system registry.
    '
    ' How to use: insert this routine in a module within your project or in a global class within
    ' your project or a referred project.
    Dim iFF%

    Dim bLog           As Boolean, bMsg As Boolean

    Static bNewSession As Boolean

    ' See if logging/msgbox is required/wanted
    Dim oAppSettings

    If (oAppSettings Is Nothing) Then
        bLog = True
        bMsg = True
    Else
        bLog = CBool(oAppSettings.GetSetting( _
            "General", "Logging", "True"))
        bMsg = CBool(oAppSettings.GetSetting( _
            "General", "ReportErrors", "False"))
    End If

    If bLog Then

        ' Logging required/wanted
        iFF = FreeFile
        Open App.Path & "\LogError.txt" For _
            Append As #iFF
        Open App.Path & "\Log.txt" For Append _
            As #iFF

        If Not bNewSession Then
            bNewSession = True
            Print #iFF, Date & "  - " & Time & _
                " --- " & _
                "New session....................................................."
        End If

        Print #iFF, Date & "  - " & Time & _
            " --- " & sErrDesc & " --- in " & _
            sLocation & " / " & Str$(iLine)
        Close #iFF
    End If

    If bMsg Then
        ' MsgBox required/wanted
        'TODO: Replace the "MyAppName" string below with your application's name
        MsgBox "Error: " & sErrDesc & vbCrLf & _
            vbCrLf & _
            "The error happened in component '" _
            & sLocation & "' at line " & Trim$( _
            iLine) & _
            " and was logged (if configured so) to the 'Log.txt' file.", _
            vbOKOnly + vbCritical, _
            "Lending System"
    End If

End Sub

Public Sub Expense()

        '<EhHeader>
        On Error GoTo Expense_Err

        '</EhHeader>

100     Set rsExpense = Nothing
102     Set rsExpense = New ADODB.Recordset

104     With rsExpense
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblExpense"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Expense_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Expense", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Loan()

        '<EhHeader>
        On Error GoTo Loan_Err

        '</EhHeader>

100     Set rsLoan = Nothing
102     Set rsLoan = New ADODB.Recordset

104     With rsLoan
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select TOP 1 * from tblLoan"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Loan_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Loan", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Loan1()

        '<EhHeader>
        On Error GoTo Loan1_Err

        '</EhHeader>

100     Set rsLoan1 = Nothing
102     Set rsLoan1 = New ADODB.Recordset

104     With rsLoan1
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select TOP 10 * from tblLoan "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Loan1_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Loan1", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Login()

        '<EhHeader>
        On Error GoTo Login_Err

        '</EhHeader>

100     Set rsLogin = Nothing
102     Set rsLogin = New ADODB.Recordset

104     With rsLogin
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblUser"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Login_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Login", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Logtime()

        '<EhHeader>
        On Error GoTo Logtime_Err

        '</EhHeader>

100     Set rsLogtime = Nothing
102     Set rsLogtime = New ADODB.Recordset

104     With rsLogtime
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblLogtime"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Logtime_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Logtime", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub payment()

        '<EhHeader>
        On Error GoTo payment_Err

        '</EhHeader>

100     Set rsPayment = Nothing
102     Set rsPayment = New ADODB.Recordset

104     With rsPayment
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select  top 1 * from tblPayment "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

payment_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.payment", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub payment1()

        '<EhHeader>
        On Error GoTo payment1_Err

        '</EhHeader>

100     Set rsPayment1 = Nothing
102     Set rsPayment1 = New ADODB.Recordset

104     With rsPayment1
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select TOP 1 * from tblPayment "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

payment1_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.payment1", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

'Public Sub paymentORnum()
'
'        '<EhHeader>
'        On Error GoTo paymentORnum_Err
'
'        TxtLog "Entered paymentORnum"
'
'        '</EhHeader>
'
'100     Set rsPaymentORnum = Nothing
'102     Set rsPaymentORnum = New adodb.Recordset
'
'104     With rsPaymentORnum
'106         .CursorType = adOpenDynamic
'108         .LockType = adLockOptimistic
'110         .ActiveConnection = conn
'112         .Source = "Select * from tblPayment "
'114         .CursorLocation = adUseClient
'116         .Open
'        End With
'
'        '<EhFooter>
'
'        TxtLog "Exited paymentORnum"
'
'        Exit Sub
'
'paymentORnum_Err:
'        ErrReport Err.Description, "LendingClientV2.DBsetup.paymentORnum", Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub
Public Sub Servicefee()

        '<EhHeader>
        On Error GoTo Servicefee_Err

        '</EhHeader>

100     Set rsServicefee = Nothing
102     Set rsServicefee = New ADODB.Recordset

104     With rsServicefee
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = _
                "Select * from tblServicefee"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

Servicefee_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.Servicefee", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Trail()
    '        '<EhHeader>
    '        On Error GoTo Trail_Err
    '
    '        '</EhHeader>
    '
    '100     Set rsTrail = Nothing
    '102     Set rsTrail = New ADODB.Recordset
    '
    '104     With rsTrail
    '106         .CursorType = adOpenDynamic
    '108         .LockType = adLockOptimistic
    '110         .ActiveConnection = conn
    '112         .Source = "Select TOP 1 * from tblTrail order by Date"
    '114         .CursorLocation = adUseClient
    '116         .Open
    '        End With
    '
    '        '<EhFooter>
    '        Exit Sub
    '
    'Trail_Err:
    '        ErrReport Err.Description, "LendingClientV2.DBsetup.Trail", Erl
    '
    '        Resume Next
    '
    '        '</EhFooter>
End Sub
'CSEH: ErrResumeNext
'Public Sub TxtLog(sText As String, Optional bNoDateTime As Boolean = False)
'
'    '<EhHeader>
'    On Error Resume Next
'
'    '</EhHeader>
'
'    ' This routine is provided to be used in conjunction with the ErrReportAndTrace _
'      error handling scheme
'
'    ' as well as for any other tasks that require logging.
'    '
'    ' How to use: insert this routine in a module within your project or in a global class within
'    ' your project or a referred project.
'    Dim iFF%, sTrailer$
'
'    Static bNewSession As Boolean
'
'    sTrailer = ""
'
'    If Not bNoDateTime Then sTrailer = Date & " - " & Time & " --- "
'    iFF = FreeFile
'    Open App.Path & "\Log.txt" For Append As #iFF
'
'    'Open App.Path & "\LogError.txt" For Append As #iFF
'    If Not bNewSession Then
'        bNewSession = True
'        Print #iFF, sTrailer & "New session....................................................."
'    End If
'
'    Print #iFF, sTrailer & sText
'    Close #iFF
'End Sub

Public Sub UnloadAllForms(Optional FormToIgnore _
    As String = "")

        '<EhHeader>
        On Error GoTo UnloadAllForms_Err

        '</EhHeader>
        Dim f As Form

100     For Each f In Forms

102         If f.Name <> FormToIgnore Then
104             Unload f
106             Set f = Nothing
            End If

108     Next f

        '<EhFooter>
        Exit Sub

UnloadAllForms_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.UnloadAllForms", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub User()

        '<EhHeader>
        On Error GoTo User_Err

        '</EhHeader>

100     Set rsUser = Nothing
102     Set rsUser = New ADODB.Recordset

104     With rsUser
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblUser"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>
        Exit Sub

User_Err:
        ErrReport Err.Description, _
            "LendingClientV2.DBsetup.User", Erl

        Resume Next

        '</EhFooter>

End Sub
