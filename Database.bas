Attribute VB_Name = "Database"
Option Explicit

Public Type ToolProcessPocketInfo
    intTOOL_REF As Integer
    lngPROCES_NR As Long
    intPocketNr As Integer
    strInfo As String
    tErrorOut As ErrorOut
End Type

Public Type LayerStrSpecInfo
    StrSpec As String
    strInfo As String
    StrLongInfo As String
End Type

Public Type UserInfo

intName_ID As Integer
intIP_ID As Integer
intHostname_ID As Integer

strHostname As String
strName As String
strIP As String

End Type

Sub text_x()
Dim tToolProcessPocketInfo As ToolProcessPocketInfo
'  "21115 010 IP",3
tToolProcessPocketInfo = GetToolProcessPocket_from_LayerName("21115 010 IP", 3)



End Sub


Public Function GetToolProcessPocket_from_LayerName(strWafer_id As String, intEpiLayerNameRef As Integer) As ToolProcessPocketInfo

    Dim objRst As New ADODB.Recordset
Dim tToolProcessPocketInfo As ToolProcessPocketInfo
    
    Set objRst = GetBatchEpiType_from_EpiLayerStrucName(intEpiLayerNameRef)


If objRst.State <> 0 Then
    If objRst.RecordCount > 0 Then
        objRst.MoveFirst



            Do Until objRst.EOF = True
            
            'r.EPI_LAYER_NAME_REF, r.GENERAL_BATCH_EPI_TYPE_EPI
            
              '  Debug.Print objRst![EPI_TYPE]
                
                tToolProcessPocketInfo = GetToolProcessPocket_from_EPI_TYPE(strWafer_id, objRst![EPI_TYPE])
                If tToolProcessPocketInfo.tErrorOut.lngErrorOut = 0 Then
                    GetToolProcessPocket_from_LayerName = tToolProcessPocketInfo
                    Exit Do
                End If
                
                objRst.MoveNext
            Loop
    Else
        GetToolProcessPocket_from_LayerName.tErrorOut.lngErrorOut = 2
        GetToolProcessPocket_from_LayerName.tErrorOut.strMessage = "Wafer ID not found"
    End If
  Else
        GetToolProcessPocket_from_LayerName.tErrorOut.lngErrorOut = 1
        GetToolProcessPocket_from_LayerName.tErrorOut.strMessage = "Status=0"
End If
    
End Function

Function GetToolProcessPocket_from_EPI_TYPE(strWafer_id As String, strEPI_TYPE As String) As ToolProcessPocketInfo

    Dim tToolProcessPocketInfo As ToolProcessPocketInfo
    Dim strSql As String
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    
    GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = -1
    
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & 3 & "'")

    strSql = "SELECT first 1 r.PROCESS_NR, r.POCKET_NR FROM FAST_EPI_OVERVIEW r where r.WAFER_NUMBER='" & strWafer_id & "' and r.EPI_STRUC='" & strEPI_TYPE & "' order by r.PROCESS_NR desc"
    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])
    
    If objRst.State <> 0 Then
        If objRst.RecordCount > 0 Then
            GetToolProcessPocket_from_EPI_TYPE.intTOOL_REF = 3
            GetToolProcessPocket_from_EPI_TYPE.lngPROCES_NR = objRst![PROCESS_NR]
            GetToolProcessPocket_from_EPI_TYPE.intPocketNr = objRst![POCKET_NR]
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 0
            Exit Function
        Else
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 2
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.strMessage = "noRecord found at SPL"
        End If
    Else
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 1
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.strMessage = "Status=0"
            Exit Function
    End If
    
    

    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & 2 & "'")

   
    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])
    
    If objRst.State <> 0 Then
        If objRst.RecordCount > 0 Then
            GetToolProcessPocket_from_EPI_TYPE.intTOOL_REF = 2
            GetToolProcessPocket_from_EPI_TYPE.lngPROCES_NR = objRst![PROCESS_NR]
            GetToolProcessPocket_from_EPI_TYPE.intPocketNr = objRst![POCKET_NR]
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 0
            Exit Function
        Else
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 2
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.strMessage = "noRecord found at SPM"
        End If
    Else
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.lngErrorOut = 1
            GetToolProcessPocket_from_EPI_TYPE.tErrorOut.strMessage = "Status=0"
            Exit Function
    End If


End Function

Function GetBatchEpiType_from_EpiLayerStrucName(intEpiLayerNameRef As Integer) As ADODB.Recordset

    Dim strSql As String
    Dim objRst As New ADODB.Recordset

    Call GlobalValues.InitGlobalValues
    strSql = "SELECT r.ID, r.EPI_LAYER_NAME_REF, r.GENERAL_BATCH_EPI_TYPE_EPI, s.EPI_TYPE FROM EPI_LAYER_STRUC_NAME r, GENERAL_BATCH_EPI_TYPE s where s.ID=r.GENERAL_BATCH_EPI_TYPE_EPI and r.EPI_LAYER_NAME_REF=" & intEpiLayerNameRef
    Set GetBatchEpiType_from_EpiLayerStrucName = Database.ReadRecords(strSql)

End Function


Public Function GetProcessFlowStep(intTOOL_REF As Integer, lngPROCES_NR As Long, intPocketNr As Integer) As Double

    Dim strSql As String
    Dim objRst As New ADODB.Recordset

    Call GlobalValues.InitGlobalValues

    GetProcessFlowStep = -1
    strSql = "SELECT s.GENRAL_TOOL_LIST_REF, s.PROCES_NR, r.POCKET_NR, r.FLOW_STEP FROM EPI_SPL_PROCES_INFO s, EPI_SPL_PROCES_WAFER_POCKET r where r.EPI_SPL_PROCES_INFO_REF=s.ID and s.GENRAL_TOOL_LIST_REF=" & intTOOL_REF & " and s.PROCES_NR=" & lngPROCES_NR & " and r.POCKET_NR=" & intPocketNr
    Set objRst = Database.ReadRecords(strSql)

    If objRst.RecordCount = 1 Then
        If Not IsNull(objRst![FLOW_STEP]) Then GetProcessFlowStep = objRst![FLOW_STEP]

    End If

End Function

Function GetMaxProcessNumber(tool_ref As Integer) As Long

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    Dim dbProcessStep As Double

    GetMaxProcessNumber = 0

    strSql = "SELECT max(r.PROCES_NR) FROM EPI_SPL_PROCES_INFO r where r.GENRAL_TOOL_LIST_REF=" & tool_ref
    Set objRst = Database.ReadRecords(strSql)

    If objRst.State <> 0 Then

        If objRst.RecordCount = 1 Then
            GetMaxProcessNumber = objRst![Max]
        End If

    Else
        Application.StatusBar = "No DB connection"
    End If

End Function

Function Get_L1_average(lngProcesNr As Long, tool_ref As Integer) As Double

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    Dim dbProcessStep As Double

    Get_L1_average = -99999
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & tool_ref & "'")

    strSql = "SELECT avg(r.L1) as L1_AVG FROM FAST_EPI_OVERVIEW r  where r.PROCESS_NR=" & lngProcesNr & ""
    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])

    If objRst.State <> 0 Then

        If objRst.RecordCount = 1 Then

            If Not IsNull(objRst![L1_AVG]) Then Get_L1_average = objRst![L1_AVG]

        End If

    Else
        Application.StatusBar = "No DB connection"

    End If

End Function

Function Get_PL_average(lngProcesNr As Long, tool_ref As Integer) As Double

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    Dim dbProcessStep As Double

    Get_PL_average = -99999
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & tool_ref & "'")

    strSql = "SELECT avg(r.PL_NM_AUTO533) as PL_AVG FROM FAST_EPI_OVERVIEW r  where r.PROCESS_NR=" & lngProcesNr & ""
    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])

    If objRst.RecordCount = 1 Then

        If Not IsNull(objRst![PL_AVG]) Then Get_PL_average = objRst![PL_AVG]

    End If

End Function

Function CreateAllDataDirectory(lngProcesNr As Long, tool_ref As Integer) As Integer

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    Dim dbProcessStep As Double

    CreateAllDataDirectory = 0
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & tool_ref & "'")

    Application.StatusBar = "Create directory"

    strSql = "SELECT r.EPI_TYPE, r.PROCESS_NR, r.POCKET_NR FROM FAST_EPI_OVERVIEW r  where r.PROCESS_NR=" & lngProcesNr & ""
    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])

    If objRst.RecordCount > 0 Then
        objRst.MoveFirst

        If objRst![EPI_TYPE] = "FINAL RUN" Then

            Do Until objRst.EOF = True
                Debug.Print CreateDataDirectory(tool_ref, objRst![PROCESS_NR], objRst![POCKET_NR], tool_ref)
                objRst.MoveNext
            Loop

        End If

    End If

    CreateAllDataDirectory = 1

End Function

Function CreateDataDirectory(lngTool_ref As Integer, lngProcesNr As Long, intPocketNr As Integer, tool_ref As Integer) As Integer

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim objRst As New ADODB.Recordset
    Dim dbProcessStep As Double
    Dim strPath As String
    Dim dStep As Double
    Dim vValue As Variant
    Dim intTmp As Integer

    strPath = "O:\03 Production data"
    CreateDataDirectory = 0
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & tool_ref & "'")

    Application.StatusBar = "Create directory"
    strSql = "SELECT r.EPI_TYPE, r.WAFER, r.WAFER_NUMBER, r.CUSTOMER, r.BATCH, r.PROCESS_NR, r.POCKET_NR FROM FAST_EPI_OVERVIEW r  where r.PROCESS_NR=" & lngProcesNr & " and r.POCKET_NR=" & intPocketNr

    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])

    If objRst.RecordCount = 1 Then
        objRst.MoveFirst

        dbProcessStep = GetProcessFlowStep(tool_ref, lngProcesNr, intPocketNr)

        If dbProcessStep = -1 Or dbProcessStep = 0 Then

            If CheckIfFolderExists(strPath) = False Then
                MsgBox "Path" & strPath & " not accesible. Mount O: drive"
                CreateDataDirectory = -2
                Exit Function
            End If

            dStep = GetProcessFlowStep(lngTool_ref, lngProcesNr, intPocketNr)

            vValue = InputBox("Current epi flow for " & lngTool_ref & "_" & lngProcesNr & "_" & intPocketNr & " step is " & dStep & "." & vbNewLine & "Give Process Flow Step")

            If IsNumeric(vValue) Then
                dbProcessStep = CDbl(vValue)
                intTmp = SetProcessFlowStep(lngTool_ref, lngProcesNr, intPocketNr, dbProcessStep)

                If intTmp = -1 Then
                    CreateDataDirectory = -1
                    Exit Function
                End If

            End If

            If objRst![EPI_TYPE] = "FINAL RUN" Then

                If IsNull(objRst![Batch]) Then
                    Application.StatusBar = "Directory not created. Batch is null"
                    CreateDataDirectory = -3
                    Exit Function
                End If

                If IsNull(objRst![customer]) Then
                    Application.StatusBar = "Directory not created. Customer is null"
                    CreateDataDirectory = -3
                    Exit Function
                End If

                Call CreateOperationDataFolder(objRst![customer], objRst![Batch], objRst![WAFER_NUMBER], dbProcessStep, , objRst![WAFER])

            Else
                CreateDataDirectory = -3
                Exit Function
            End If

        End If

    End If

End Function

Function CreateOperationDataFolder(strSP As String, strBatch As String, strWaferID As String, dFlowStep As Double, Optional strPath As String = "O:\03 Production data", Optional strMsg As String) As Integer

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Dim strPathCreate As String

    CreateOperationDataFolder = 0

    Call GlobalValues.InitGlobalValues

    strPathCreate = strPath

    If CheckIfFolderExists(strPathCreate) = False Then
        MsgBox "Path" & strPath & " not accesible. Mount O: drive"
        CreateOperationDataFolder = -1
        Exit Function
    End If

    strPathCreate = strPathCreate & "\" & strSP

    If CheckIfFolderExists(strPathCreate) = False Then
        Call MakeMyFolder(strPathCreate)

        If CheckIfFolderExists(strPathCreate) = False Then
            CreateOperationDataFolder = -2
            Exit Function
        End If

    End If

    strPathCreate = strPathCreate & "\" & strBatch

    If CheckIfFolderExists(strPathCreate) = False Then
        Call MakeMyFolder(strPathCreate)

        If CheckIfFolderExists(strPathCreate) = False Then
            CreateOperationDataFolder = -2
            Exit Function
        End If

    End If

    strPathCreate = strPathCreate & "\" & strWaferID

    If CheckIfFolderExists(strPathCreate) = False Then
        Call MakeMyFolder(strPathCreate)

        If CheckIfFolderExists(strPathCreate) = False Then
            CreateOperationDataFolder = -2
            Exit Function
        End If

    End If

    strPathCreate = strPathCreate & "\" & dFlowStep

    If CheckIfFolderExists(strPathCreate) = False Then
        Call MakeMyFolder(strPathCreate)

        If CheckIfFolderExists(strPathCreate) = False Then
            CreateOperationDataFolder = -2
            Exit Function
        End If

    End If

    CreateOperationDataFolder = 1

    Call CreateTextFileMsg(strWaferID, strPathCreate & "\" & strMsg & ".txt")

    If CheckIfFolderExists(strPathCreate & "\" & strWaferID) = False Then
        CreateOperationDataFolder = -3
        Exit Function
    End If

    CreateOperationDataFolder = 2

End Function

Public Function GetWaferID(intTOOL_REF As Integer, lngPROCES_NR As Long, intPocketNr As Integer) As String

    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    GetWaferID = ""
    Call GlobalValues.InitGlobalValues
    bDBonline = IsSQLRunning(strCnn)

    strSql = "SELECT first 1 r.WAFER_NUMBER FROM FAST_EPI_OVERVIEW r where r.PROCESS_NR=" & lngPROCES_NR & " and r.POCKET_NR=" & intPocketNr

    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & intTOOL_REF & "'")

    Set objRst = Database.ReadRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])

    If objRst.RecordCount = 1 Then
        If Not IsNull(objRst![WAFER_NUMBER]) Then GetWaferID = objRst![WAFER_NUMBER]
    End If

End Function

Public Function SetProcessFlowStep(intTOOL_REF As Integer, lngPROCES_NR As Long, intPocketNr As Integer, dFlowStep As Double) As Integer

    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Dim lngID As Long
    SetProcessFlowStep = 0
    Call GlobalValues.InitGlobalValues
    bDBonline = IsSQLRunning(strCnn)

    strSql = "SELECT r.ID, r.BATCH FROM EPI_SPL_PROCES_INFO s, EPI_SPL_PROCES_WAFER_POCKET r where r.EPI_SPL_PROCES_INFO_REF=s.ID and s.GENRAL_TOOL_LIST_REF=" & intTOOL_REF & " and s.PROCES_NR=" & lngPROCES_NR & " and r.POCKET_NR=" & intPocketNr
    Set objRst = Database.ReadRecords(strSql)

    If objRst.RecordCount = 1 Then
        If Not IsNull(objRst![ID]) Then Debug.Print "Test: " & objRst![ID]

        strSql = "UPDATE EPI_SPL_PROCES_WAFER_POCKET r set r.FLOW_STEP=" & dFlowStep & "  where r.ID=" & objRst![ID]
        Call Database.UpdateRecords(strSql)
        SetProcessFlowStep = 1

    Else
        SetProcessFlowStep = -1
        Exit Function
    End If

End Function

Public Function GetSubEpiLayerLast(lngGENERAL_CUSTOMER_BATCH_Ref As Long) As Integer

    Dim strSql As String
    Dim objRst As New ADODB.Recordset

    GetSubEpiLayerLast = -1
    strSql = "SELECT COALESCE(max(r.SUB_EPI_ORDER),0) as LayerLast FROM GENERAL_BATCH_EPISTACK r where r.GENERAL_BATCH_REF=" & lngGENERAL_CUSTOMER_BATCH_Ref
    Set objRst = Database.ReadRecords(strSql)

    If objRst.RecordCount = 1 Then
        GetSubEpiLayerLast = objRst![LayerLast]
    End If

End Function

Sub UPDATE_FAST_EPI_OVERVIEW(Optional lngProcesNr As Long = 0, Optional intPocketNr As Integer = 0, Optional intFASTREFRESH As Integer = 0, Optional tool_ref As Integer = 3)

    Dim strSql As String
    Dim dblSecondsElapsed As Double
    Dim dblStartTime As Double
    Dim tErrorOut As ErrorOut
    Call GlobalValues.InitGlobalValues
    recGENERAL_PARAMETERS.MoveFirst
    recGENERAL_PARAMETERS.Find ("search like 'strCnn_bulk_" & tool_ref & "'")

    Application.StatusBar = "EXECUTE PROCEDURE"
    strSql = "EXECUTE PROCEDURE UPDATE_FAST_EPI_OVERVIEW_3X ('" & lngProcesNr & "'); "
    tErrorOut = Database.UpdateRecordsErrOut(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])
    Call StdErrorDBHandling(tErrorOut, recGENERAL_PARAMETERS![PARAM_VALUE], strSql)
    dblStartTime = Timer

    Application.StatusBar = "Finished: " & Round(Timer - dblStartTime, 2) & " " & Replace(strSql, "EXECUTE PROCEDURE", "")
    Application.Wait (Now + TimeValue("0:00:2"))
    Application.StatusBar = "commit;"
    strSql = "commit;"
    'Call ReadRecords(strSql)
    Call Database.UpdateRecords(strSql, recGENERAL_PARAMETERS![PARAM_VALUE])
End Sub

Function ReadRecords(strSql As String, Optional strCnn_new As String = strCnn) As ADODB.Recordset
    Dim objCnn As New ADODB.Connection
    Dim objRst As New ADODB.Recordset

    On Error GoTo errorhandling
    objCnn.Open strCnn_new
    objRst.Open strSql, objCnn, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Set ReadRecords = objRst

    Application.StatusBar = "objRst.State: " & " -> " & objRst.State

    Set objRst = Nothing
    Exit Function
errorhandling:
    Application.StatusBar = Err.Description & " -> " & Err.Number
    On Error GoTo 0
    Err.Clear

End Function

Function ReadRecordsNew(strSql As String, Optional strCnn_new As String = strCnn) As AdobeErrorOut
    Dim objCnn As New ADODB.Connection
    Dim objRst As New ADODB.Recordset

    On Error GoTo errorhandling
    objCnn.Open strCnn_new
    objRst.Open strSql, objCnn, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Set ReadRecordsNew.objRstOut = objRst

    Application.StatusBar = "objRst.State: " & " -> " & objRst.State

    Set objRst = Nothing
    Exit Function
errorhandling:
    Application.StatusBar = Err.Description & " -> " & Err.Number
    On Error GoTo 0
    Err.Clear

End Function

Function ReadRecordsGlobal(strSql As String) As ADODB.Recordset

    'If Not objCnnGlobal Is Nothing Then
    '    If objCnnGlobal.State = adStateClose Then
    '        objCnnGlobal.Open strCnn
    '        Debug.Print objCnnGlobal.State, adStateOpen
    '    End If
    'End If

    'On Error GoTo errorhandling

    Debug.Print objCnnGlobal.State, adStateOpen
    If objCnnGlobal.State = adStateOpen Then objCnnGlobal.Close
    objRstGlobal.Open strSql, objCnnGlobal, adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Set ReadRecordsGlobal = objRstGlobal

    'Set objRst = Nothing
    Exit Function

errorhandling:
    Application.StatusBar = "Database is busy. Please repeat last operation"

End Function

Function IsSQLRunning(sCN$) As Boolean
    On Error Resume Next
    Dim Conn As ADODB.Connection

    Err.Clear
    Set Conn = New ADODB.Connection
    Conn.Open sCN

    If Err = 0 Then
        IsSQLRunning = True
    End If

    Conn.Close
    Set Conn = Nothing
End Function

Public Function UpdateRecordsErrOut(strSql As String, Optional strCnn_new As String = strCnn) As ErrorOut
    Dim objCnn As New ADODB.Connection
    Dim objRst As New ADODB.Recordset

    On Error GoTo errorhandling

    Call GlobalValues.InitGlobalValues

    If bDBonline = True Then
        objCnn.Open strCnn_new

        '  Debug.Print strSql

        'Debug.Print strSql
        objRst.Open strSql, objCnn, adOpenStatic
        objCnn.Close
    Else
        Application.StatusBar = "No Connection to DB. bDBonline=False."
    End If

    UpdateRecordsErrOut.dTime = Now()
    UpdateRecordsErrOut.lngErrorOut = Err.Number
    UpdateRecordsErrOut.strMessage = ""
    Exit Function
errorhandling:
    UpdateRecordsErrOut.dTime = Now()
    UpdateRecordsErrOut.lngErrorOut = Err.Number
    UpdateRecordsErrOut.strMessage = Err.Description
    Err.Clear
End Function

Public Function UpdateRecords(strSql As String, Optional strCnn_new As String = strCnn) As ErrorOut
    Dim objCnn As New ADODB.Connection
    Dim objRst As New ADODB.Recordset
    Err.Clear
    UpdateRecords.lngErrorOut = 0
    On Error GoTo errorhandling

    objCnn.Open strCnn_new

    Debug.Print objCnn.State

    If objCnn.State > 0 Then


        objRst.Open strSql, objCnn, adOpenStatic
        objCnn.Close

        UpdateRecords.lngErrorOut = Err.Number
        UpdateRecords.strMessage = Err.Description
        UpdateRecords.strFunction = strSql
        Exit Function
    Else

        Application.StatusBar = "objCnn.State " & 0
        UpdateRecords.lngErrorOut = 0
        UpdateRecords.strMessage = "objCnn.State=0"
        UpdateRecords.strFunction = strSql
        Exit Function

    End If

errorhandling:

    Application.StatusBar = "UpdateRecords: " & Err.Number & " " & Err.Description
    UpdateRecords.lngErrorOut = Err.Number
    UpdateRecords.strMessage = Err.Description
    UpdateRecords.strFunction = strSql

    Err.Clear
    On Error GoTo 0
End Function

Public Function GetUpdateUserData() As UserInfo
    Dim strSql As String
    Dim strName As String
    Dim strHostname As String
    Dim intHostname_ID As Integer
    Dim intName_ID As Integer
    Dim strIP As String
    Dim strIP_ID As Integer
    Dim objCnn As New ADODB.Connection
    Dim objRst As New ADODB.Recordset
    Dim strTMP() As String
    Call GlobalValues.InitGlobalValues
    strHostname = GeneralLib.Hostname()
    strName = GeneralLib.Username()

    GetUpdateUserData.strHostname = strHostname
    GetUpdateUserData.strName = strName

    strSql = "UPDATE or INSERT INTO GENERAL_USER_LIST (USER_NAME) VALUES ('" & strName & "') matching(USER_NAME)"
    Call UpdateRecords(strSql)

    strSql = "UPDATE or INSERT INTO GENERAL_HOSTNAME_LIST (HOSTNAME) VALUES ('" & strHostname & "') matching(HOSTNAME)"
    Call UpdateRecords(strSql)
    Set objRst = Nothing
    strSql = "select rdb$get_context('SYSTEM', 'CLIENT_ADDRESS') as CLIENT_IP from rdb$database"
    Set objRst = ReadRecords(strSql)

    If bDBonline = True Then
        strIP = objRst![CLIENT_IP]
        strTMP() = Split(strIP, "/")
        strIP = strTMP(0)

        GetUpdateUserData.strIP = strIP

        strSql = "UPDATE or INSERT INTO GENERAL_CLIENT_IP_LIST (IP) VALUES ('" & strIP & "') matching(IP)"
        Call UpdateRecords(strSql)

        strSql = "SELECT r.ID FROM GENERAL_CLIENT_IP_LIST r where r.IP='" & strIP & "'"
        Set objRst = ReadRecords(strSql)
        GetUpdateUserData.intIP_ID = objRst![ID]

        strSql = "SELECT r.ID FROM GENERAL_USER_LIST r where r.USER_NAME='" & strName & "'"
        Set objRst = ReadRecords(strSql)
        GetUpdateUserData.intName_ID = objRst![ID]

        strSql = "SELECT r.ID FROM GENERAL_HOSTNAME_LIST r where r.HOSTNAME='" & strHostname & "'"
        Set objRst = ReadRecords(strSql)
        GetUpdateUserData.intHostname_ID = objRst![ID]
    End If

    Set objRst = Nothing
End Function

Public Function GetServerStatus(intID As Integer) As String
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    strSql = "SELECT r.ID, r.REQUEST, r.COMMAND,r.TOOL_NAME_REF, r.COMMAND_TIME, r.STATUS, r.STATUS_UPDATE_TIME , r.UPDATE_TIME, r.INFO, r.PRIO, r.FOLDER_MOD_TIME FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set objRst = ReadRecords(strSql)
    GetServerStatus = objRst![Status]
End Function

Public Function GetServerCommand(intID As Integer) As String
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    strSql = "SELECT r.ID, r.REQUEST, r.COMMAND, r.COMMAND_TIME, r.STATUS, r.STATUS_UPDATE_TIME , r.UPDATE_TIME, r.INFO, r.PRIO, r.FOLDER_MOD_TIME FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set objRst = ReadRecords(strSql)
    GetServerCommand = ""

    If Not IsNull(objRst![Command]) Then GetServerCommand = objRst![Command]
End Function

Public Function GetToolRef(intID As Integer) As String
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    strSql = "SELECT r.TOOL_NAME_REF FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set objRst = ReadRecords(strSql)
    GetToolRef = objRst![TOOL_NAME_REF]
End Function

Public Function GetStatusRaw(intID As Integer) As ADODB.Recordset
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    strSql = "SELECT * FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set GetStatusRaw = ReadRecords(strSql)
End Function

Public Function GetDataFolder(intID As Integer) As String
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    strSql = "SELECT r.DATA_FOLDER FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set objRst = ReadRecords(strSql)
    GetDataFolder = objRst![DATA_FOLDER]
End Function

Public Function GetFOLDER_MOD_TIME(intID As Integer) As String
    Dim strSql As String
    Dim objRst As New ADODB.Recordset
    Call GlobalValues.InitGlobalValues
    bDBonline = IsSQLRunning(strCnn)
    strSql = "SELECT r.ID, r.REQUEST, r.COMMAND, r.COMMAND_TIME, r.STATUS, r.STATUS_UPDATE_TIME , r.UPDATE_TIME, r.INFO, r.PRIO, r.FOLDER_MOD_TIME FROM SERVER_STATUS_UPDATE r where ID=" & intID
    Set objRst = ReadRecords(strSql)
    GetFOLDER_MOD_TIME = 0

    If Not IsNull(objRst![FOLDER_MOD_TIME]) Then GetFOLDER_MOD_TIME = objRst![FOLDER_MOD_TIME]

End Function

Public Function SetFOLDER_MOD_TIME(dtDate As Date, intID As Integer)
    Dim strSql As String

    strSql = "UPDATE SERVER_STATUS_UPDATE a SET a.FOLDER_MOD_TIME = '" & Format(dtDate, "dd.MM.yyyy, HH:mm:ss") & "' WHERE a.ID = " & intID
    Call Database.UpdateRecords(strSql)

End Function

Public Function SetServerStatus(strStatus As String, intID As Integer)
    Dim strSql As String

    strSql = "UPDATE SERVER_STATUS_UPDATE a SET a.STATUS = '" & Left(Replace(strStatus, "'", "`"), 300) & "' WHERE a.ID = " & intID
    Call Database.UpdateRecords(strSql)

End Function

Public Function SetServerInfo(strInfo As String, intID As Integer)
    Dim strSql As String

    strSql = "UPDATE SERVER_STATUS_UPDATE a SET a.info = '" & strInfo & "' WHERE a.ID = " & intID
    Call Database.UpdateRecords(strSql)

End Function

Public Function SetServerCommand(strCommand As String, intID As Integer)
    Dim strSql As String

    strSql = "UPDATE SERVER_STATUS_UPDATE a SET a.COMMAND = '" & strCommand & "' WHERE a.ID = " & intID
    Call Database.UpdateRecords(strSql)

End Function

Sub testPL()
Dim tErrOutLast As ErrorOut
tErrOutLast = UploadPLData(3)
End Sub

Sub RefreshAllFastTable()

    Dim strSql As String
    Dim strStatus As String
    Dim statusID As Integer
    Dim tErrOutLast As ErrorOut
    Dim tErrorOut As ErrorOut
    Dim intErrorTotal As Integer
    Dim i As Integer
    Dim strPath As String
    Dim objStatusRaw As ADODB.Recordset

    intErrorTotal = 0
StartEngine:
    On Error GoTo errorhandler
    bDBonline = IsSQLRunning(strCnn)
    Call InitGlobalValues

    If bDBonline = True Then

        Do
        bDBonline = IsSQLRunning(strCnn)

        If bDBonline = True Then

            Call SetServerStatus("CheckModificationTimes", 1)
            
            Call CheckModificationTimes

            Call SetServerStatus("CheckModificationTimes finished", 1)

            '#################################################################################################

            statusID = 2
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Call SetServerCommand("", statusID)
                Call SetServerStatus("Running ReBuildTableAll", statusID)

                Call ReBuildTableAll(3)
                Call SetServerStatus("Running ImportUsingQueryTable", statusID)
                Call ImportUsingQueryTable(100000, 3)
                Call SetServerStatus("Writing down data", statusID)
                Call WriteCSVs
                Call SetServerStatus("Finished", statusID)

            End If

            '#################################################################################################

            statusID = 3
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then
                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running UploadPLData", statusID)
                    tErrOutLast = UploadPLData(statusID)
                    Call SetServerStatus("Finished ErrOut=" & tErrOutLast.lngErrorOut & "  Msg=" & tErrOutLast.strMessage, statusID)
                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)

            End If

            '#################################################################################################

            statusID = 4
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    tErrOutLast = UploadParticlesData(statusID)
                    Call SetServerStatus("Finished ErrOut=" & tErrOutLast.lngErrorOut & "  Msg=" & tErrOutLast.strMessage, statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)

            End If

            '#################################################################################################

            statusID = 5
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then
                Call SetServerCommand("", statusID)
                tErrOutLast = LoopThroughEpiFiles(statusID)
                Call SetServerStatus("Finished ErrOut=" & tErrOutLast.lngErrorOut & "  Msg=" & tErrOutLast.strMessage, statusID)
            End If

            ' upload cac data

            statusID = 7
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    tErrorOut = UploadEpiCacFilesAll(statusID)
                    If tErrorOut.lngErrorOut <> 0 Then
                        Call SetServerInfo(tErrorOut.strMessage & " Err=" & tErrorOut.lngErrorOut, statusID)
                    Else
                        Call SetServerInfo("Ok Err=" & tErrorOut.lngErrorOut, statusID)
                    End If
                    
                    Call SetServerStatus("Finished", statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################
            
                        statusID = 17
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call UploadEpiCacFilesAll(statusID)
                    Call SetServerStatus("Finished", statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################

            statusID = 9
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call UploadCVdata(statusID)
                    Call SetServerStatus("Finished", statusID)
                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################

            statusID = 18
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call UploadCVdata(statusID)
                    Call SetServerStatus("Finished", statusID)
                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################
            'upload L-reactor LOG files
            '#################################################################################################
            
            statusID = 11
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    tErrOutLast = UploadEpiLogFile(statusID)
                    Call SetServerStatus("Finished", statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################

            statusID = 12
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call UploadCVdata(statusID)
                    Call SetServerStatus("Finished", statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)
            End If

            '#################################################################################################

            statusID = 13
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then
                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call UploadParticlesData(statusID)
                    Call SetServerStatus("Finished", statusID)
                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)

            End If

            '#################################################################################################

            statusID = 14
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running", statusID)
                    Call LoopThroughEpiFiles(statusID)
                    Call SetServerStatus("Finished", statusID)

                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)

            End If

            '#################################################################################################

            statusID = 15
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                If CheckIfFolderExists(strPath) Then

                    Call SetServerCommand("", statusID)
                    Call SetServerStatus("Running UploadPLData", statusID)
                    Call UploadPLData(statusID)
                    Call SetServerStatus("Finished", statusID)
                Else
                    Call SetServerStatus("Not exists: " & strPath, statusID)
                End If

                Call SetServerCommand("", statusID)

            End If

            '#################################################################################################

            statusID = 16
            strStatus = GetServerCommand(statusID)

            If strStatus = "ON" Then
                Call SetServerCommand("", statusID)
                Call SetServerStatus("Running ReBuildTableAll", statusID)
                Call ReBuildTableAll(2)
                Call SetServerStatus("Running ImportUsingQueryTable", statusID)
                Call ImportUsingQueryTable(100000, 2)
                Call SetServerStatus("Writing down data", statusID)
                Call WriteCSVs
                Call SetServerStatus("Finished", statusID)

            End If

            Call SetServerStatus("Coffee pause 90sec", 1)
        Else

            Call Pause(90, "Server not found waiting ...")

        End If

        '#################################################################################################

        Call SetServerStatus("Coffee pause 90sec", 1)
        Call Pause(90)

    Loop Until GetServerCommand(1) <> "ON"

    Call SetServerStatus("Stopped", 1)
    Call SetServerCommand("", 1)

    Exit Sub
    Else
        MsgBox "Server not found"
    End If

    ' this is the only way to stop engine

    Exit Sub

errorhandler:
        Call SetServerStatus("Error StatusID@" & statusID & " Err: " & Err.Number & " Desc:" & Err.Description, 1)
        Call Pause(200)
        bDBonline = IsSQLRunning(strCnn)
        intErrorTotal = intErrorTotal + 1

        If intErrorTotal > 200 Then Exit Sub

            If bDBonline = True Then
                tErrOutLast.strMessage = tErrOutLast.strMessage & "RefreshAllFastTable LastStatusID@" & statusID & " Err.Description: " & Err.Description
                tErrOutLast.lngErrorOut = Err.Number
                tErrOutLast.dTime = Now()

                Debug.Print tErrOutLast.strMessage
                Call UploadErrorOut(tErrOutLast)

                For i = 2 To i < 16
                    Call SetServerCommand("", i)
                    Call SetServerStatus("Reset", i)
                    Call SetFOLDER_MOD_TIME("01.01.1900, 01:01:01", i)
                Next i

            End If

            Err.Clear
            GoTo StartEngine
        End Sub

        Public Sub Pause(sngSecs As Single, Optional strMsg As String = "Coffe time 90s:  ")
            Dim sngEnd As Single
            sngEnd = Timer + sngSecs

            While Timer < sngEnd
            DoEvents
            Application.StatusBar = strMsg & (sngEnd - Timer)

            If (sngEnd - Timer) > 50000 Then Exit Sub
                Wend

            End Sub

            Sub CheckModificationTimes()
                Dim fso
                Dim oFolder
                Dim oSubfolder
                Dim oFile
                Dim queue As Collection
                Dim strFile() As String
                Dim strPath As String
                Dim UserInfoAct As UserInfo
                Dim strFileModTime As String
                Dim dtDate As Date
                Dim dtDateDB As Date
                Dim statusID As Integer
                Dim objStatusRaw As ADODB.Recordset
                Dim lngTool_ref As Long
                'Check PL
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set queue = New Collection

                UserInfoAct = GetUpdateUserData()

                dtDate = 0
                statusID = 9

                Call InitGlobalValues
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking CV data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPL: New CV data found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPL: No new CV data", statusID)
                End If

                dtDate = 0
                statusID = 12

                Call InitGlobalValues
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking CV data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPL: New CV data found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPL: No new CV data", statusID)
                End If

                dtDate = 0
                statusID = 3

                Call InitGlobalValues
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking PL data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPL: New PL data found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPL: No new PL data", statusID)
                End If

                dtDate = 0
                statusID = 15

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking PL data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPM: New PL data found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPM: No new PL data", statusID)
                End If

                dtDate = 0
                statusID = 18

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking PL data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("M0306_CV: New data found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("M0306_CV: No New data found", statusID)
                End If



                dtDate = 0
                statusID = 4

                Call InitGlobalValues
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking PL data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPL: New particle file found", statusID)
                    Call SetServerCommand("ON", statusID)
                Else
                    Call SetServerStatus("SPL: no new particle file found", statusID)
                End If

                dtDate = 0

                statusID = 13

                Call InitGlobalValues
                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking PL data", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then

                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPM: New particle file found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPM: no new particle file found", statusID)
                End If

                dtDate = 0
                statusID = 5

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking Epi files", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then
                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPL: New Epi files found", statusID)
                    Call SetServerCommand("ON", statusID)

                Else
                    Call SetServerStatus("SPL: No new epi files", statusID)
                End If

                dtDate = 0
                statusID = 14

                Set objStatusRaw = GetStatusRaw(statusID)
                strPath = objStatusRaw![DATA_FOLDER]

                Call SetServerStatus("Looking Epi files", statusID)
                dtDateDB = GetFOLDER_MOD_TIME(statusID)
                dtDate = GetLatestFolderModificationDate(strPath)

                If dtDateDB <> dtDate Then
                    Call SetFOLDER_MOD_TIME(dtDate, statusID)
                    Call SetServerStatus("SPM: New Epi files found", statusID)
                    Call SetServerCommand("ON", statusID)
                Else
                    Call SetServerStatus("SPM: No new epi files", statusID)
                End If

            End Sub

            Function GetLatestFolderModificationDate(strPath As String) As Date
                Dim fso
                Dim oFolder
                Dim oSubfolder
                Dim oFile
                Dim queue As Collection
                Dim strFile() As String

                Dim UserInfoAct As UserInfo
                Dim strFileModTime As String
                Dim dtDate As Date
                'Check PL
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set queue = New Collection

                dtDate = 0
                queue.Add fso.GetFolder(strPath) 'obviously replace

                Do While queue.Count > 0
                    Set oFolder = queue(1)
                    queue.Remove 1 'dequeue
                    '...insert any folder processing code here...

                    For Each oSubfolder In oFolder.subfolders
                        queue.Add oSubfolder 'enqueue
                    Next oSubfolder

                    For Each oFile In oFolder.Files

                        strFileModTime = Format(FileDateTime(oFile), "dd.MM.yyyy, HH:mm:ss")

                        If dtDate < oFile.DateLastModified Then
                            dtDate = oFile.DateLastModified

                        End If

                    Next oFile

                Loop

                GetLatestFolderModificationDate = dtDate

            End Function

            
