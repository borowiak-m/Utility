Function ConnectionString(env)
    ' Function sets a connection string based on chosen environment passed as argument
    If LCase(env) = "prod" then
        ConnectionString=   "Provider=MSOLEDBSQL;Server=" & ProdSQLServerName & ";" & _
		    			    "Integrated Security=SSPI;Database=" & ProdDatabaseName & ";" & _
		    				"DataTypeCompatibility=80;MultiSubnetFailover=True"
    Elseif LCase(env) = "test" then
        ConnectionString=   "Provider=MSOLEDBSQL;Server=" & TestSQLServerName & ";" & _
		    			    "Integrated Security=SSPI;Database=" & TestDatabaseName & ";" & _
		    				"DataTypeCompatibility=80;MultiSubnetFailover=True"
    Else
        ConnectionString = False 
    End If
End Function

' Add order to timed out orders table
Function InsertTimedoutOrder(ordnum)
    On Error Resume Next
    Dim obj_SQLConnection:    Set obj_SQLConnection     = CreateObject("ADODB.Connection")
    Dim obj_SQLCommand:       Set obj_SQLCommand        = CreateObject("ADODB.Command")
    Dim objParamOrdnum1:      Set objParamOrdnum1       = CreateObject("ADODB.Parameter")
    Dim objParamOrdnum2:      Set objParamOrdnum2       = CreateObject("ADODB.Parameter")
    Dim objParamOrdnum3:      Set objParamOrdnum3       = CreateObject("ADODB.Parameter")
    Dim str_SQLQuery:         str_SQLQuery              = ""
    Dim str_FunctionName:     str_FunctionName          = "InsertTimedoutOrder"

    obj_SQLConnection.Open ConnectionString(currentEnv)

    str_SQLQuery = "IF EXISTS (SELECT 1 FROM ordersplitting.TimedoutOrders WHERE ordnum = ?) " & _
                   "BEGIN " & _
                   "    UPDATE ordersplitting.TimedoutOrders " & _
                   "    SET createdts = GETDATE() " & _
                   "    WHERE ordnum = ?; " & _
                   "END " & _
                   "ELSE " & _
                   "BEGIN " & _
                   "    INSERT INTO ordersplitting.TimedoutOrders (ordnum) " & _
                   "    VALUES (?); " & _
                   "END"

    With obj_SQLCommand
        .ActiveConnection = obj_SQLConnection
        .CommandText = str_SQLQuery
        .CommandType = 1 'adCmdText

        Set objParamOrdnum1 = .CreateParameter("ordnum1", 3, 1)
        Set objParamOrdnum2 = .CreateParameter("ordnum2", 3, 1)
        Set objParamOrdnum3 = .CreateParameter("ordnum3", 3, 1)

        .Parameters.Append objParamOrdnum1
        objParamOrdnum1.Value = ordnum
        .Parameters.Append objParamOrdnum2
        objParamOrdnum2.Value = ordnum
        .Parameters.Append objParamOrdnum3
        objParamOrdnum3.Value = ordnum
    End With

    obj_SQLCommand.Execute
    if Err.Number <> 0 then Error_handler(str_FunctionName)
    Set obj_SQLConnection = Nothing
    Set obj_SQLCommand = Nothing
End Function

Function GetTimedoutOrderNumbers(minutesTimeout)
    On Error Resume Next

    Dim obj_SQLConnection:          Set obj_SQLConnection     = CreateObject("ADODB.Connection")
    Dim obj_RecordSet:              Set obj_RecordSet         = CreateObject("ADODB.Recordset")
    Dim obj_SQLCommand:             Set obj_SQLCommand        = CreateObject("ADODB.Command")
    Dim objParamHours:              Set objParamHours         = CreateObject("ADODB.Parameter")
    Dim str_SQLQuery:               str_SQLQuery              = ""
    Dim str_FunctionName:           str_FunctionName          = "GetTimedoutOrderNumbers"
    Dim str_ConcatenatedOrderNums:  str_ConcatenatedOrderNums = ""
    GetTimedoutOrderNumbers = ""
    obj_SQLConnection.Open ConnectionString(currentEnv)

    str_SQLQuery = "select distinct ordnum from ordersplitting.TimedoutOrders where createdts >= DATEADD(mi, -?, getdate())"

    If (not IsNumeric(minutesTimeout)) then 
        minutesTimeout = 1
    Else 
        If (minutesTimeout<0) then minutesTimeout = -minutesTimeout
    End If

    With obj_SQLCommand
        .ActiveConnection = obj_SQLConnection
        .CommandText = str_SQLQuery
        .CommandType = 1 'adCmdText

        Set objParamHours = .CreateParameter("timeout", 3, 1)
        .Parameters.Append objParamHours
        objParamHours.Value = minutesTimeout
    End With

    obj_RecordSet.Open obj_SQLCommand, , adOpenStatic, adLockReadOnly
    if Err.Number <> 0 then Error_handler(str_FunctionName)

    If Not obj_RecordSet.EOF Then 
        Call ConcatenateOrderNums(obj_RecordSet, str_ConcatenatedOrderNums)
    End If

    obj_RecordSet.Close
    Set obj_RecordSet = Nothing
    obj_SQLConnection.Close
    Set obj_SQLConnection = Nothing
    Set obj_SQLCommand = Nothing

    GetTimedoutOrderNumbers = str_ConcatenatedOrderNums
End Function
