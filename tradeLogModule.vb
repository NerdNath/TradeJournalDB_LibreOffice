REM Global variables
Dim globalDlg			'run-time dialog that is created
Dim activeOrderIndex	'a variable defining what order is selected in order table
Dim actvExeIndex		'a variable defining what execution is selected in table
Dim masterExeList(0,0)	'variable holding execution data & relating it to ordr indexes
Dim editField			'a variable for the edit text field service name string
Dim numField			'a variable for the numeric field service name string
Dim timeField			'a variable for the time field service name string
Dim curField			'a variable for the currency field service name string
Dim combBox				'a variable for the combo box field service name string
Dim dateField			'a variable for the date field service name string

Sub runLogDlg
	REM Declare primary variables
	Dim tradeLogLib 	'Library that contains the logTradeDlg dialog
	Dim tradeLogDlg		'Dialog as it is stored in logTradeLib
	Dim baseDlgModel	'The model to the tradeLogDlg
	
	
	REM Load library and dialog
	DialogLibraries.loadLibrary("tradeLogLib")
	tradeLogLib = DialogLibraries.getByName("tradeLogLib")
	tradeLogDlg = tradeLogLib.getByName("logTradeDlg")
	
	REM initialize global variables
	globalDlg = CreateUnoDialog(tradeLogDlg)
	activeOrderIndex = "none"
	actvExeIndex = "none"

	editField = "stardiv.vcl.controlmodel.Edit"
	numField = "stardiv.vcl.controlmodel.NumericField"
	timeField = "stardiv.vcl.controlmodel.TimeField"
	curField = "stardiv.vcl.controlmodel.CurrencyField"
	combBox = "stardiv.vcl.controlmodel.ComboBox"
	dateField = "stardiv.vcl.controlmodel.DateField"

	REM Store the model to the tradeLogDlg in a variable for future
	baseDlgModel = globalDlg.getmodel()
	
	REM Build the data model for orders, and insert the display table
	buildInsertOrders(baseDlgModel)
	
	REM Build the data model for executions, and insert the display table
	buildInsertExes(baseDlgModel)
	
	REM Run the dialog
	globalDlg.execute()
	'inspeCommand
	'debugFunc
End Sub

'Debugging tools and scripts follow

Function debugFunc
	Dim i%
	For i% = LBound(globalDlg.Model.ControlModels) to UBound(globalDlg.Model.ControlModels)
		Print "[" + ObjToString(i%) +"] " + globalDlg.Model.ControlModels(i%).Name
	Next
End Function

Function inspeCommand
	Dim tblList
	Dim colList
	tblList = ThisDatabaseDocument.CurrentController.ActiveConnection.getTables()
	colList = tblList(0).getColumns()
	'Print globalDlg.Model.ControlModels(2).getServiceName()
	'Inspect(globalDlg.GetControl("ordersTbl"))
	'Inspect(globalDlg.getByIdentifier(53).Model)
	'Inspect(globalDlg.Model.ControlModels(27))
	Inspect(ThisDatabaseDocument.CurrentController.ActiveConnection.CreateStatement())
	'Inspect(ThisDatabaseDocument.CurrentController.DataSource.Tables.GetByIndex(2).getIdentifier())
End Function

Function entrOrdPress
    Dim ctrlModels
	Dim ordTlbDataModel
    Dim firstFieldIndex
    Dim lastFieldIndex
    Dim orderArray(7)
    Dim i

    'store the indexes of the 1st order field, last order field, & 1st execution field
    firstFieldIndex = 18
    lastFieldIndex = 25

    ctrlModels = globalDlg.Model.ControlModels
    ordTlbDataModel = ctrlModels(0).GridDataModel

    'While the index is one of the order fields, check for empty fields
    'If empty field is encountered, set focus to it and exit the function
	'collect the contents of the orders fields & store them in an array
    'also disable the fields so they can't be edited further at this point
    i = firstFieldIndex
    Do While (i <= lastFieldIndex)
		Select Case ctrlModels(i).getServiceName()
			Case numField, curField
				If IsEmpty(ctrlModels(i).Value) = True Then
					globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
					Exit Function
				End If
				orderArray(i-firstFieldIndex) = ctrlModels(i).Value
			Case combBox
				If ctrlModels(i).Text = "" Then
					globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
					Exit Function
				End If
				orderArray(i-firstFieldIndex) = ctrlModels(i).Text
			Case dateField
				If IsEmpty(ctrlModels(i).Date) = True Then
					globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
					Exit Function
				End If
				orderArray(i-firstFieldIndex) = ctrlModels(i).Date
			Case Else
				Print "There's a problem"
				Exit Function
		End Select
		ctrlModels(i).Enabled = False
        i = i + 1
    Loop

    ordTlbDataModel.addRow("",orderArray) 'add the order row into the table
	activeOrderIndex = ordTlbDataModel.RowCount - 1
    globalDlg.GetControl("ordersTbl").selectRow(activeOrderIndex) 'select the row
	'set activeOrderIndex to pending value so that the executions will be held unti
	'"finish order" is clicked
    activeOrderIndex = "pend"
    ctrlModels(lastFieldIndex + 1).Enabled = False  'disable the button
    globalDlg.GetControl("exeTimeField").SetFocus()
End Function

Function nextExePress
    Dim ctrlModels
    Dim exeTblDataModel
    Dim firstFieldIndex
    Dim lastFieldIndex
    Dim exeArray(2)
    Dim i%

    'store the indexes of 1st exe field and last exe field
    firstFieldIndex = 27
    lastFieldIndex = 29

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

    If activeOrderIndex = "none" Then Exit Function
    'TODO find where focus is and put it back there
    'basically make this button not accept focus unless it was tabbed to
    'before it was clicked, aka, it already had focus before this function ran

    'While the index is one of the order fields, do work on that field:
    'If empty field is encountered, set focus to it and exit the function
	'collect the contents of the execution fields and store them in exeArray
	'clear all the execution fields and make them enabled
    i = firstFieldIndex
    Do While (i <= lastFieldIndex)
		Select Case ctrlModels(i).getServiceName()
			Case numField, curField
				If IsEmpty(ctrlModels(i).Value) = True Then
					globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
					Exit Function
				End If
				exeArray(i - firstFieldIndex) = ctrlModels(i).Value
				ctrlModels(i).Value = Empty
			Case timeField
				If ctrlModels(i).Text = "" Then
					globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
					Exit Function
				End If
				exeArray(i - firstFieldIndex) = ctrlModels(i).Text
				ctrlModels(i).Text = ""
			Case Else
				Print "There's a problem"
				Exit Function
		End Select
        i = i + 1
    Loop

    'Place the contents of the fields in the execution table data model
    exeTblDataModel.addRow("",exeArray)

    'Return everything to the "enter new execution" state
	actvExeIndex = "none"
	globalDlg.GetControl("exeTable").deselectAllRows()
    globalDlg.GetControl(ctrlModels(firstFieldIndex).Name).SetFocus()
	If activeOrderIndex = "pend" Then
		Exit Function
	Else
		injectNewExe(exeArray)
	End If
End Function 

Function exeUpClick
    Dim ctrlModels      'array of all control models in the dialog
    Dim exeTblDataModel 'the data model for the executions table
    Dim firstFieldIndex
    Dim lastFieldIndex
    Dim i

    'store indexes of 1st exe field and last exe field
    firstFieldIndex = 27
    lastFieldIndex = 29

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

    If activeOrderIndex = "none" Then Exit Function
	If exeTblDataModel.RowCount = 0 Then Exit Function

    'if active exe = "none" then set it to last row that is part of active order
    'else if there is an execution before it, set it to that
    'else exit function
    If actvExeIndex = "none" Then
        actvExeIndex = exeTblDataModel.RowCount -1
    ElseIf (actvExeIndex > 0) Then
        actvExeIndex = actvExeIndex - 1
    Else
        Exit Function
    End If

    globalDlg.GetControl("exeTable").deselectAllRows()
    globalDlg.GetControl("exeTable").selectRow(actvExeIndex)

    'fill the boxes with the contents of active execution then disable the boxes
    For i = firstFieldIndex To lastFieldIndex
		Select Case ctrlModels(i).getServiceName()
			Case numField, curField
				ctrlModels(i).Value = _
                     exeTblDataModel.getCellData(i - firstFieldIndex, actvExeIndex)
			Case timeField
				ctrlModels(i).Text = _
                     exeTblDataModel.getCellData(i - firstFieldIndex, actvExeIndex)
			Case Else
				Print "There's a problem"
				Exit Function
		End Select
        ctrlModels(i).Enabled = False
    Next
	ctrlModels(30).Enabled = False 'disable the "Next Execution" button
End Function

Function exeDwnClick
    Dim ctrlModels      'array of all control models in the dialog
    Dim exeTblDataModel 'the data model for the executions table
    Dim firstFieldIndex
    Dim lastFieldIndex
    Dim i

    'store indexes of the 1st exe field and last exe field
    firstFieldIndex = 27
    lastFieldIndex = 29

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

    If activeOrderIndex = "none" Then Exit Function

    If actvExeIndex = "none" Then
        Exit Function
    ElseIf actvExeIndex = exeTblDataModel.RowCount - 1 Then
        actvExeIndex = "none"
        For i = firstFieldIndex to lastFieldIndex
			Select Case ctrlModels(i).getServiceName()
				Case numField, curField
            		ctrlModels(i).Value = Empty
				Case timeField
					ctrlModels(i).Text = ""
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
            ctrlModels(i).Enabled = True
        Next
        globalDlg.GetControl("exeTable").deselectAllRows()
		ctrlModels(30).Enabled = True 'enable the "Next Execution" button
        globalDlg.GetControl(ctrlModels(firstFieldIndex).Name).SetFocus()
    Else
        actvExeIndex = actvExeIndex + 1
        globalDlg.GetControl("exeTable").deselectAllRows()
        globalDlg.GetControl("exeTable").selectRow(actvExeIndex)
        For i = firstFieldIndex To lastFieldIndex
			Select Case ctrlModels(i).getServiceName()
				Case numField, curField
            		ctrlModels(i).Value = _
                        exeTblDataModel.getCellData(i - firstFieldIndex, actvExeIndex)
				Case timeField
					ctrlModels(i).Text = _
                        exeTblDataModel.getCellData(i - firstFieldIndex, actvExeIndex)
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
        Next
    End If
End Function

Function delExeBtnClick
    Dim ctrlModels
    Dim exeTblDataModel
    Dim tempArray
    Dim i

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

    If activeOrderIndex = "none" Then
        Exit Function
    ElseIf actvExeIndex = "none" Then
        Exit Function
	ElseIf activeOrderIndex = "pend" Then
		exeTblDataModel.removeRow(actvExeIndex)
		actvExeIndex = "none"
		selectActvExecution()
		Exit Function
    ElseIf exeTblDataModel.RowCount = 1 Then
        masterExeList(activeOrderIndex,0) = empty
        exeTblDataModel.removeAllRows()
        selectActvExecution()
    ElseIf actvExeIndex = exeTblDataModel.RowCount - 1 Then
        masterExeList(activeOrderIndex,actvExeIndex) = empty
        exeTblDataModel.removeRow(actvExeIndex)
        actvExeIndex = "none"
        selectActvExecution()
    ElseIf actvExeIndex = 0 Then
        Redim tempArray(0 To Ubound(masterExeList,2) - 1)
        For i = 1 To Ubound(masterExeList,2)
            tempArray(i - 1) = masterExeList(activeOrderIndex,i)
        Next
        For i = 0 To Ubound(masterExeList,2)
            masterExeList(activeOrderIndex,i) = empty
        Next
        For i = 0 To Ubound(tempArray)
            masterExeList(activeOrderIndex,i) = tempArray(i)
        Next
        exeTblDataModel.removeRow(actvExeIndex)
        selectActvExecution()
    Else
        ReDim tempArray(0 To Ubound(masterExeList,2) - 1)
        For i = 0 To actvExeIndex - 1
            tempArray(i) = masterExeList(activeOrderIndex,i)
        Next
        For i = actvExeIndex + 1 To Ubound(masterExeList,2)
            tempArray(i - 1) = masterExeList(activeOrderIndex,i)
        Next
        For i = 0 To Ubound(masterExeList,2)
            masterExeList(activeOrderIndex,i) = empty
        Next
        For i = 0 To Ubound(tempArray)
            masterExeList(activeOrderIndex,i) = tempArray(i)
        Next
        exeTblDataModel.removeRow(actvExeIndex)
        selectActvExecution()
    End If
End Function

Function finishOrdPress
    Dim ctrlModels
    Dim exeTblDataModel
    Dim i

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

	If activeOrderIndex = "none" Then Exit Function

	'If the finish order button was clickable, then the last row of the orders table
	'should have been a new order and activeOrderIndex should be "pend" until now
	activeOrderIndex = ctrlModels(0).GridDataModel.RowCount - 1

    'store all the values in the executions table into a multidimensional array that
    'reflects the relationship of the active order's index, and those values
    If exeTblDataModel.RowCount - 1 > UBound(masterExeList,2) Then
        Redim Preserve masterExeList(0 To activeOrderIndex, 0 To _
                                     exeTblDataModel.RowCount -1)
    Else
        Redim Preserve masterExeList(0 To activeOrderIndex, 0 To _
                                    UBound(masterExeList,2))
    End If

    For i = 0 To exeTblDataModel.RowCount - 1
        masterExeList(activeOrderIndex, i) = exeTblDataModel.getRowData(i)
    Next

    'set the active order and active execution global index variables to "none"
    activeOrderIndex = "none"
	existingOrderSelected()
End Function

Function ordrUpClick
    Dim ordGC           'orders grid control
    Dim ctrlModels
    Dim ordTlbDataModel

    ctrlModels = globalDlg.Model.ControlModels
    ordTlbDataModel = ctrlModels(0).GridDataModel

    'If the orders table has nothing in it, exit function
    If ordTlbDataModel.RowCount = 0 Then Exit Function

    'Decide what to set the activeOrderIndex to based on what it currently is
    If activeOrderIndex = "none" Then activeOrderIndex = ordTlbDataModel.RowCount
	If activeOrderIndex = "pend" Then Exit Function
    If activeOrderIndex > 0 Then
         activeOrderIndex = activeOrderIndex - 1
		 existingOrderSelected()
    ElseIf activeOrderIndex = 0 Then
        Exit Function
    End If
End Function

Function ordrDownClick
    Dim ordGC           'orders grid control
    Dim ctrlModels
    Dim ordTlbDataModel
    
    ctrlModels = globalDlg.Model.ControlModels
    ordTlbDataModel = ctrlModels(0).GridDataModel

    'If the orders table has nothing in it, exit function
    If ordTlbDataModel.RowCount = 0 Then Exit Function

    If activeOrderIndex = "none" Then
        Exit Function
	ElseIf activeOrderIndex = "pend" Then
		Exit Function
    ElseIf activeOrderIndex = ordTlbDataModel.RowCount - 1 Then
        activeOrderIndex = "none"
        existingOrderSelected()
    Else
        activeOrderIndex = activeOrderIndex + 1
        existingOrderSelected()
    End If
End Function


Function delOrdClick
    Dim ctrlModels
    Dim ordTlbDataModel
    Dim tempArray
    Dim i
    Dim x

    ctrlModels = globalDlg.Model.ControlModels
    ordTlbDataModel = ctrlModels(0).GridDataModel

    If activeOrderIndex = "none" Then
        Exit Function
	ElseIf activeOrderIndex = "pend" Then
		activeOrderIndex = "none"
		ordTlbDataModel.removeRow(ordTlbDataModel.RowCount - 1)
		existingOrderSelected()
    ElseIf ordTlbDataModel.RowCount = 1 Then
        activeOrderIndex = "none"
        Redim masterExeList(0,0)
        ordTlbDataModel.removeAllRows()
        existingOrderSelected()
    ElseIf activeOrderIndex = ordTlbDataModel.RowCount - 1 Then
        activeOrderIndex = "none"
        Redim Preserve masterExeList(0 To UBound(masterExeList,1) - 1, _
                                     0 To UBound(masterExeList,2))
        ordTlbDataModel.removeRow(ordTlbDataModel.RowCount - 1)
        existingOrderSelected()
    ElseIf activeOrderIndex = 0 Then
        Redim tempArray(0 To UBound(masterExeList,1) - 1, _
                        0 To UBound(masterExeList,2))
        For i = 1 To UBound(masterExeList,1)
            For x = 0 To UBound(masterExeList,2)
                tempArray(i - 1,x) = masterExeList(i,x)
            Next
        Next
        Redim masterExeList(0 To UBound(tempArray,1), 0 To UBound(tempArray,2))
        masterExeList = tempArray
        ordTlbDataModel.removeRow(0)
        existingOrderSelected()
    Else
        Redim tempArray(0 To UBound(masterExeList,1) - 1, _
                        0 To UBound(masterExeList,2))
        For i = 0 To activeOrderIndex - 1
            For x = 0 To UBound(masterExeList,2)
                tempArray(i,x) = masterExeList(i,x)
            Next
        Next
        For i = activeOrderIndex + 1 To UBound(masterExeList,1)
            For x = 0 To UBound(masterExeList,2)
                tempArray(i,x) = masterExeList(i,x)
            Next
        Next
        Redim masterExeList(0 To UBound(tempArray,1), 0 To UBound(tempArray,2))
        masterExeList = tempArray
        ordTlbDataModel.removeRow(activeOrderIndex)
        existingOrderSelected()
    End If
End Function

Function existingOrderSelected
    Dim ctrlModels
    Dim ordTlbDataModel
    Dim exeTblDataModel
    Dim i

    ctrlModels = globalDlg.Model.ControlModels
    ordTlbDataModel = ctrlModels(0).GridDataModel
    exeTblDataModel = ctrlModels(1).GridDataModel

    If activeOrderIndex = "none" Then
        For i = 18 To 31 ' 18 is actNumField, 31 is finishOrderBtn
            ctrlModels(i).Enabled = True
        Next
        For i = 18 To 25
			Select Case ctrlModels(i).getServiceName()
				Case numField,curField
					ctrlModels(i).Value = Empty
				Case combBox
            		ctrlModels(i).Text = ""
				Case dateField
					ctrlModels(i).Date = Empty
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
        Next
        For i = 27 To 29
			Select Case ctrlModels(i).getServiceName()
				Case numField,curField
					ctrlModels(i).Value = Empty
				Case timeField
            		ctrlModels(i).Text = ""
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
        Next
        actvExeIndex = "none"
        exeTblDataModel.removeAllRows()
        globalDlg.getControl("ordersTbl").deselectAllRows()
        globalDlg.getControl(ctrlModels(18).Name).SetFocus()
    Else
        For i = 18 To 31
            ctrlModels(i).Enabled = False
        Next
        For i = 18 To 25
			Select Case ctrlModels(i).getServiceName()
				Case numField,curField
					ctrlModels(i).Value = ordTlbDataModel.getCellData(i - 18, _
																	 activeOrderIndex)
				Case combBox
            		ctrlModels(i).Text = ordTlbDataModel.getCellData(i - 18, _
																	 activeOrderIndex)
				Case dateField
					ctrlModels(i).Date = ordTlbDataModel.getCellData(i - 18, _
																	 activeOrderIndex)
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
        Next
        globalDlg.getControl("ordersTbl").deselectAllRows()
        globalDlg.getControl("ordersTbl").selectRow(activeOrderIndex)
        exeTblDataModel.removeAllRows()
        For i = LBound(masterExeList,2) To UBound(masterExeList,2)
            If IsEmpty(masterExeList(activeOrderIndex,i)) = False Then
                exeTblDataModel.addRow("",masterExeList(activeOrderIndex,i))
            End If
        Next
        If exeTblDataModel.RowCount > 0 Then
        	actvExeIndex = exeTblDataModel.RowCount - 1
			globalDlg.getControl("exeTable").deselectAllRows()
			globalDlg.getControl("exeTable").selectRow(actvExeIndex)
			For i = 27 To 29
				Select Case ctrlModels(i).getServiceName()
					Case numField,curField
						ctrlModels(i).Value = exeTblDataModel.getCellData(i - 27, _ 
																		actvExeIndex)
					Case timeField
						ctrlModels(i).Text = exeTblDataModel.getCellData(i - 27, _ 
																		actvExeIndex)
					Case Else
						Print "There's a problem"
						Exit Function
				End Select
        	Next
        Else
        	actvExeIndex = "none"
			For i = 27 To 29
				Select Case ctrlModels(i).getServiceName()
					Case numField,curField
						ctrlModels(i).Value = Empty
					Case timeField
						ctrlModels(i).Text = ""
					Case Else
						Print "There's a problem"
						Exit Function
				End Select
				ctrlModels(i).Enabled = True
        	Next
			ctrlModels(30).Enabled = True
        End If
    End If
End Function

Function selectActvExecution
    Dim ctrlModels
    Dim exeTblDataModel
    Dim i

    ctrlModels = globalDlg.Model.ControlModels
    exeTblDataModel = ctrlModels(1).GridDataModel

    If actvExeIndex = "none" Then
        For i = 27 To 30
            ctrlModels(i).Enabled = True
        Next
        For i = 27 To 29
			Select Case ctrlModels(i).getServiceName()
				Case numField,curField
					ctrlModels(i).Value = Empty
				Case timeField
            		ctrlModels(i).Text = ""
				Case Else
					Print "There's a problem"
					Exit Function
			End Select
        Next
        globalDlg.getControl("exeTable").deselectAllRows()
        globalDlg.getControl("exeTimeField").SetFocus()
    ElseIf exeTblDataModel.RowCount > 0 Then
        For i = 27 To 30
            ctrlModels(i).Enabled = False
        Next
        For i = 27 To 29
            Select Case ctrlModels(i).getServiceName()
					Case numField,curField
						ctrlModels(i).Value = exeTblDataModel.getCellData(i - 27, _ 
																		actvExeIndex)
					Case timeField
						ctrlModels(i).Text = exeTblDataModel.getCellData(i - 27, _ 
																		actvExeIndex)
					Case Else
						Print "There's a problem"
						Exit Function
				End Select
        Next
        globalDlg.getControl("exeTable").deselectAllRows()
        globalDlg.getControl("exeTable").selectRow(actvExeIndex)
    Else
        actvExeIndex = "none"
        selectActvExecution()
    End If
End Function

Function injectNewExe(executionArray)
    Dim exeTblDataModel

    exeTblDataModel = globalDlg.Model.ControlModels(1).GridDataModel
    
    If Ubound(masterExeList,2) < exeTblDataModel.RowCount - 1 Then
        Redim Preserve masterExeList(0 To UBound(masterExeList,1), _
                                        0 To Ubound(masterExeList,2) + 1)
    End If
    masterExeList(activeOrderIndex,Ubound(masterExeList,2)) = executionArray
End Function

Function checkTradeClosed As Boolean
	Dim buySum
	Dim sellSum
	Dim ordTlbDataModel
	Dim i
	Dim x

	buySum = 0
	sellSum = 0
	ordTlbDataModel = globalDlg.Model.ControlModels(0).GridDataModel

	For i = Lbound(masterExeList,1) To Ubound(masterExeList,1)
		For x = Lbound(masterExeList,2) To Ubound(masterExeList,2)
			If IsEmpty(masterExeList(i,x)) = False Then
				If ordTlbDataModel.getCellData(2,i) = "BUY" Then
					buySum = buySum + masterExeList(i,x)(1)
				ElseIf ordTlbDataModel.getCellData(2,i) = "SELL" Then
					sellSum = sellSum + masterExeList(i,x)(1)
				Else
					Print "There's a problem"
				End If
			End If
		Next
	Next

	checkTradeClosed = CBool(buySum = sellSum)
	If ordTlbDataModel.RowCount < 1 Then checkTradeClosed = False
End Function

Function recordTrade
	Dim dataSource
	Dim sqlStatement
	Dim ctrlModels
	Dim ordTlbDataModel
	Dim exeTblDataModel
	Dim queryResult
	Dim tradeTotal
	Dim s$
	Dim i
	Dim x

	dataSource = ThisDatabaseDocument.CurrentController
	ctrlModels = globalDlg.Model.ControlModels
	ordTlbDataModel = globalDlg.Model.ControlModels(0).GridDataModel
	exeTblDataModel = globalDlg.Model.ControlModels(1).GridDataModel

	If IsNull(dataSource.ActiveConnection) Then dataSource.connect
	sqlStatement = dataSource.ActiveConnection.CreateStatement()
	s$ = "SELECT COUNT(*) FROM ""TradeList"""
	queryResult = sqlStatement.executeQuery(s$)
	queryResult.next()
	tradeTotal = queryResult.getInt(1)
	If tradeTotal = 0 Then tradeTotal = -1

	If checkTradeClosed() = True Then
		s$ = "INSERT INTO ""TradeList""(pattern_name,ticker," + _
					"exchange,trade_type,approx_float,SSS_PP,SSS_RR,SSS_EE," + _
					"SSS_PF,SSS_TS,SSS_RC,SSS_ME,q1,q2,q3,q4) VALUES ("
		For i = 2 To 17
			Select Case ctrlModels(i).getServiceName()
				Case editField,combBox
					If ctrlModels(i).Text = "" Then
						globalDlg.GetControl(ctrlModels(i).Name).SetFocus()
						Exit Function
					Else
						s$ = s$ + "'" + ctrlModels(i).Text + "',"
					End If
				Case numField
					If IsEmpty(ctrlModels(i).Value) = True Then
						globalDlg.getControl(ctrlModels(i).Name).SetFocus()
						Exit Function
					Else
						s$ = s$ + Str(ctrlModels(i).Value) + ","
					End If
			End Select
		Next
		s$ = Left(s$,Len(s$) - 1)	'trims off the comma at the end
		s$ = s$ + ");"
		sqlStatement.execute(s$)
		s$ = ""

		For i = 0 To ordTlbDataModel.RowCount - 1
			s$ = "INSERT INTO ""Orders""(trade_id,account_number,order_number," + _
  					"action,submit_date,exe_type,requested_price,status," + _
  					"qty_requested) VALUES ("
			s$ = s$ + str(tradeTotal + 1) + ","
			For x = 0 To 7
				Select Case x
					Case 0,1,5,7
						s$ = s$ + str(ordTlbDataModel.getCellData(x,i)) + ","
					Case 2,4,6
						s$ = s$ + "'" + ordTlbDataModel.getCellData(x,i) + "',"
					Case 3
						s$ = s$ +"'"+ Trim(str(ordTlbDataModel.getCellData(x,i).Year))
						s$ = s$ +"-"+ Trim(str(ordTlbDataModel.getCellData(x,i).Month))
						s$ = s$ +"-"+ Trim(str(ordTlbDataModel.getCellData(x,i).Day))
						s$ = s$ + "',"
				End Select
			Next
			s$ = Left(s$,Len(s$) - 1) 'trims off the comma at the end
			s$ = s$ + ");"
			sqlStatement.execute(s$)
			s$ = ""
		Next

		For i = Lbound(masterExeList,1) To Ubound(masterExeList,1)
			For x = Lbound(masterExeList,2) To Ubound(masterExeList,2)
				If IsEmpty(masterExeList(i,x)) = False Then
					s$ = "INSERT INTO ""Executions""(trade_id,account_number," + _
  							"order_number,time_of_execution,quantity_executed," + _ 
							"price_executed) VALUES ("
					s$ = s$ + str(tradeTotal + 1) + ","
					s$ = s$ + str(ordTlbDataModel.getCellData(0,i)) + ","
					s$ = s$ + str(ordTlbDataModel.getCellData(1,i)) + ","
					s$ = s$ + "'" + masterExeList(i,x)(0) + "',"
					s$ = s$ + str(masterExeList(i,x)(1)) + ","
					s$ = s$ + str(masterExeList(i,x)(2)) + ");"
					sqlStatement.execute(s$)
					s$ = ""
				End If
			Next
		Next
		clearForm()
		Print "Trade Recorded Successfully"
	Else
		Print "Trade is open, check executions"
	End If
End Function

Function clearForm
	Dim i

	masterExeList = Empty
	globalDlg.Model.ControlModels(0).GridDataModel.removeAllRows()
	globalDlg.Model.ControlModels(1).GridDataModel.removeAllRows()

	For i = 2 To 25
		Select Case globalDlg.Model.ControlModels(i).getServiceName()
			Case editField,combBox
				globalDlg.Model.ControlModels(i).Text = ""
			Case curField,numField
				globalDlg.Model.ControlModels(i).Value = Empty
			Case dateField
				globalDlg.Model.ControlModels(i).Date = Empty
			Case Else
				Print "There's a problem"
				Exit Function
		End Select
	Next

	For i = firstFieldIndex to lastFieldIndex
		Select Case globalDlg.Model.ControlModels(i).getServiceName()
			Case numField, curField
				globalDlg.Model.ControlModels(i).Value = Empty
			Case timeField
				globalDlg.Model.ControlModels(i).Text = ""
			Case Else
				Print "There's a problem"
				Exit Function
		End Select
	Next
End Function

Function buildInsertOrders(parentDlgModel)
	Dim ordersGridM 	'The grid model for the orders table
    Dim ordersColM      'The column model for the orders table
    Dim ordersHeaders   'Becomes array to hold column headers for orders
    Dim ordersColWidths	'Becomes array to hold column widths for orders
    Dim col             'A variable for adding and editing columns to models
    Dim i%              'For loop increment variable

	'create new grid control and set its initial values
	ordersGridM = parentDlgModel.createInstance(_
        "com.sun.star.awt.grid.UnoControlGridModel")
    With ordersGridM
        .PositionX = 205
        .PositionY = 60
        .Width = 321
        .Height = 110
        .UseGridLines = True
        .SelectionModel = com.sun.star.view.SelectionType.MULTI
    End With

    'add name and columns
	ordersGridM.Name = "ordersTbl"
    ordersColM = ordersGridM.ColumnModel
    ordersHeaders = Array(	"Account Number",_
							"Order Number",_
							"Action",_
							"Submit Date",_
							"Execution Type",_
							"Requested Price",_
							"Status",_
							"Quantity Requested"_
						  )
    For i = LBound(ordersHeaders) to UBound(ordersHeaders)
		col = ordersColM.createColumn()
		col.Title = ordersHeaders(i)
		ordersColM.addColumn(col)
	Next

    'insert the grid control to the dialog
	parentDlgModel.insertByName("ordersTbl",ordersGridM)
	
	'Set Column style values
	ordersColWidths = Array(50,40,20,40,45,50,20,55)
	For i = LBound(ordersColWidths) to UBound(ordersColWidths)
		col = ordersColM.GetColumn(i)
		col.ColumnWidth = ordersColWidths(i)
		col.HorizontalAlign = 1
		col.Resizeable = False
	Next
End Function

Function buildInsertExes(parentDlgModel)
	Dim exeGridM		'The grid model for the executions table
    Dim exeColM			'The column model for the executions table
    Dim exeHeaders		'Becomes array to hold column headers for executions
    Dim exeColWidths	'Becomes array to hold column widths for executions
    Dim col             'A variable for adding and editing columns to models
    Dim i%              'For loop increment variable

	'create new grid control and set it's initial values
	exeGridM = parentDlgModel.createInstance(_
		"com.sun.star.awt.grid.UnoControlGridModel")
	With exeGridM
		.PositionX = 540
		.PositionY = 60
		.Width = 141
		.Height = 110
		.UseGridLines = True
		.SelectionModel = com.sun.star.view.SelectionType.MULTI
	End With

    'add name and columns
	exeGridM.Name = "exeTable"
	exeColM = exeGridM.ColumnModel
	exeHeaders = Array(	"Execution Time",_
						"Quantity Executed",_
						"Price Executed"_
					   )
	For i = LBound(exeHeaders) to UBound(exeHeaders)
		col = exeColM.createColumn()
		col.Title = exeHeaders(i)
		exeColM.addColumn(col)
	Next

    'insert the grid control to the dialog
	parentDlgModel.insertByName("exeTable",exeGridM)
	
	'set column style values
	exeColWidths = Array(45,55,40)
	For i = LBound(exeColWidths) to UBound(exeColWidths)
		col = exeColM.GetColumn(i)
		col.ColumnWidth = exeColWidths(i)
		col.HorizontalAlign = 1
		col.Resizeable = False
	Next
End Function

'Listing 544. Identify and remove white space from a string
REM Utility Functions and Methods

'*************************************************************************
'** Is the specified character whitespace? The answer is true if the
'** character is a tab, CR, LF, space, or a non-breaking space character!
'** These correspond to the ASCII values 9, 10, 13, 32, and 160
'*************************************************************************
'Function IsWhiteSpace(iChar As Integer) As Boolean
'	Select Case iChar
'	Case 9, 10, 13, 32, 160
'		IsWhiteSpace = True
'	Case Else
'		IsWhiteSpace = False
'	End Select
'End Function

'*************************************************************************
'** Find the first character starting at location i% that is not whitespace.
'** If there are none, then the return value will be greater than the
'** length of the string.
'*************************************************************************
Function FirstNonWhiteSpace(ByVal i%, s$) As Integer
	If i <= Len(s) Then
		Do While IsWhiteSpace(Asc(Mid$(s, i, 1)))
			i = i + 1
			If i > Len(s) Then
				Exit Do
			End If
		Loop
	End If
	FirstNonWhiteSpace = i
End Function

'*************************************************************************
'** Remove white space text from both the front and the end of a string.
'** This modifies the argument string and it returns the modified string.
'** This removes all types of white space, not just a regular space.
'*************************************************************************
Function TrimWhite(s As String) As String
		s = Trim(s)
	Do While Len(s) > 0
		If Not IsWhiteSpace(ASC(s)) Then Exit Do
		s = Right(s, Len(s) - 1)
	Loop
	Do While Len(s) > 0
		If Not IsWhiteSpace(ASC(Right(s, 1))) Then Exit Do
		s = Left(s, Len(s) - 1)
	Loop
	TrimWhite = s
End Function

'Listing 545. String representation of an object
Function ObjToString(oInsObj, optional arraySep$, _
			Optional maxDepth%, Optional CRReplace As Boolean) As String

	Dim iMaxDepth% ' Maximum number of array elements to display
	Dim sArraySep$ ' Separator used for array elements
	Dim s$ ' Contains the results.
	Dim bCRReplace As Boolean

	' Optional arguments
	If IsMissing(CRReplace) Then
		bCRReplace = False
	Else
		bCRReplace = CRReplace
	End If

	If IsArray(oInsObj) Then
		If IsMissing(arraySep) Then
			sArraySep = CHR$(10)
		Else
			sArraySep = arraySep
		End If
	
		If IsMissing(maxDepth) THen
			iMaxDepth = 100
		Else
			iMaxDepth = maxDepth
		End If

		s = ObjArrayToString(oInsObj, sArraySep, iMaxDepth)
	Else
		' Not an array
		Select Case VarType(oInsObj)
			Case 0
				s = "Empty"
			Case 1
				s = "Null"
			Case 2, 3, 4, 5, 7, 8, 11, 16 To 23, 33 To 37
				s = CStr(oInsObj)
			Case Else
				If IsUnoStruct(oInsObj) Then
					s = "[Cannot convert an UnoStruct to a string]"
				Else
					Dim sImplName$ : sImplName = GetImplementationName(oInsObj)
					If sImplName = "" Then
						s = "[Cannot convert " & TypeName(oInsObj) & " to a string]"
					Else
						s = "[Cannot convert " & TypeName(oInsObj) & _
							" (" & sImplName & ") to a string]"
					End If
			End If
		End Select
		If bCRReplace Then
			s = Replace(s, CHR$(10), "|", 1, -1, 1)
			s = Replace(s, CHR$(13), "|", 1, -1, 1)
		End If
		If Len(s) > 100 Then s = Left(s, 100) & "...."
	End If
	ObjToString = s
End Function

'Listing 546. Convert an array to a string
'*************************************************************************
'** Convert an array of objects to a string if possible.
'**
'** oInsObj - Object to Inspect
'** sArraySep - Separator character for arrays
'** maxDepth - Maximum number of array elements to show
'*************************************************************************
Function ObjArrayToString(oInsObj, sArraySep$, Optional maxDepth%) As String
	Dim iMaxDepth%
	Dim iDimensions%

	If IsMissing(maxDepth) THen
		iMaxDepth = 100
	Else
		iMaxDepth = maxDepth
	End If

	iDimensions = NumArrayDimensions(oInsObj)
	If iDimensions = 1 Then
		ObjArrayToString = ObjArray_1_ToString(oInsObj, iMaxDepth%, sArraySep)
	ElseIf iDimensions = 2 Then
		ObjArrayToString = ObjArray_2_ToString(oInsObj, iMaxDepth%, sArraySep)
	Else
		ObjArrayToString = "[Unable to convert array with dimension " & iDimensions & "]"
	End If
End Function

'*************************************************************************
'** Convert a 1-D array of objects to a string.
'**
'** oInsObj - Object to Inspect
'** iMaxDepth - Maximum number of array elements to show
'** sSep - Separator character for arrays
'*************************************************************************
Function ObjArray_1_ToString(oInsObj, iMaxDepth%, sSep$) As String
	Dim iDim_1%
	Dim iMax_1%
	Dim iCounter%
	Dim s$
	Dim sTempSep$
	
	sTempSep = ""
	iDim_1 = LBound(oInsObj, 1)
	iMax_1 = UBound(oInsObj, 1)
	Do While (iCounter < iMaxDepth AND iDim_1 <= iMax_1)
		iCounter = iCounter + 1
		s = s & sTempSep & "[" & iDim_1 & "] " _
			  & ObjToString(oInsObj(iDim_1), ", ")
		sTempSep = sSep
		iDim_1 = iDim_1 + 1
	Loop
	ObjArray_1_ToString = s
End Function

'*************************************************************************
'** Convert a 2-D array of objects to a string.
'**
'** oInsObj - Object to Inspect
'** iMaxDepth - Maximum number of array elements to show
'** sSep - Separator character for arrays
'*************************************************************************
Function ObjArray_2_ToString(oInsObj, iMaxDepth%, sSep$) As String
	Dim iDim_1%, iDim_2%
	Dim iMax_1%, iMax_2%
	Dim iCounter%
	Dim s$
	Dim sTempSep$

	sTempSep = ""
	iDim_1 = LBound(oInsObj, 1)
	iMax_1 = UBound(oInsObj, 1)
	iMax_2 = UBound(oInsObj, 2)
	Do While (iCounter < iMaxDepth AND iDim_1 <= iMax_1)
		iCounter = iCounter + 1
		iDim_2 = LBound(oInsObj, 2)
		Do While (iCounter < iMaxDepth AND iDim_2 <= iMax_2)
			s = s & sTempSep & "[" & iDim_1 & ", " & iDim_2 & "] " _
				  & ObjToString(oInsObj(iDim_1, iDim_2), ", ")
			sTempSep = sSep
			iDim_2 = iDim_2 + 1
		Loop
		iDim_1 = iDim_1 + 1
	Loop
	ObjArray_2_ToString = s
End Function

'*************************************************************************
'** Number of dimension for an array. a(4) returns 1 and a(3, 4) returns 2.
'**
'** oInsObj - Object to Inspect
'*************************************************************************
Function NumArrayDimensions(oInsObj) As Integer
	Dim i% : i = 0
	If IsArray(oInsObj) Then
		On Local Error Goto DebugBoundsError:
		Do While (i% >= 0)
			LBound(oInsObj, i% + 1)
			UBound(oInsObj, i% + 1)
			i% = i% + 1
		Loop
		DebugBoundsError:
		On Local Error Goto 0
	End If
	NumArrayDimensions = i
End Function

'Listing 547. Get the object's implementation name (if it is an UNO Service)
'*************************************************************************
'** Get an objects implementation name if it exists.
'**
'** oInsObj - Object to Inspect
'*************************************************************************
Function GetImplementationName(oInsObj) As String
	On Local Error GoTo DebugNoSet
	Dim oTmpObj
	oTmpObj = oInsObj
	GetImplementationName = oTmpObj.getImplementationName()
	DebugNoSet:
End Function

'Listing 548. Get information about an object
'*************************************************************************
'** Return a string that contains useful information about an object.
'** This includes detailed information about the object type. If the
'** object is an array, then the array dimensions are listed.
'** If possible, even the UNO implementation name is obtained.
'*************************************************************************
Function ObjInfoString(oInsObj) As String
	Dim s As String

	REM We can always get the type name and variable type.
	s = "TypeName = " & TypeName(oInsObj) & CHR$(10) &_
	"VarType = " & VarType(oInsObj) & CHR$(10)

	REM Check for NULL and EMPTY
	If IsNull(oInsObj) Then
		s = s & "IsNull = True"
	ElseIf IsEmpty(oInsObj) Then
		s = s & "IsEmpty = True"
	Else
		If IsObject(oInsObj) Then
			s = s & "Implementation Name = " & GetImplementationName(oInsObj)
			s = s & CHR$(10) & "IsObject = True" & CHR$(10)
		End If
		If IsUnoStruct(oInsObj) Then s = s & "IsUnoStruct = True" & CHR$(10)
		If IsDate(oInsObj) Then s = s & "IsDate = True" & CHR$(10)
		If IsNumeric(oInsObj) Then s = s & "IsNumeric = True" & CHR$(10)
		If IsArray(oInsObj) Then
			On Local Error Goto DebugBoundsError:
			Dim i%, sTemp$
			s = s & "IsArray = True" & CHR$(10) & "range = ("
			Do While (i% >= 0)
				i% = i% + 1
				sTemp$ = LBound(oInsObj, i%) & " To " & UBound(oInsObj, i%)
				If i% > 1 Then s = s & ", "
				s = s & sTemp$
			Loop
			DebugBoundsError:
			On Local Error Goto 0
			s = s & ")" & CHR$(10)
		End If
	End If

	s = s & "Value : " & CHR$(10) & ObjToString(oInsObj) & CHR$(10)
	ObjInfoString = s
End Function

'Listing 549. Indirectly Sort an array by moving indexes
'*************************************************************************
'** Sort the oItems() array by arranging the iIdx() array.
'** If oItems() = ("B", "C", "A") Then on output, iIdx() = (2, 0, 1)
'** This means that the item that should come first is oItems(2), second
'** is oItems(0), and finally, oItems(1)
'*************************************************************************
Sub SortMyArray(oItems(), iIdx() As Integer)
	Dim i As Integer 'Outer index variable
	Dim j As Integer 'Inner index variable
	Dim temp As Integer 'Temporary variable to swap two values.
	Dim bChanged As Boolean 'Becomes True when something changes
	For i = LBound(oItems()) To UBound(oItems()) - 1
		bChanged = False
		For j = UBound(oItems()) To i+1 Step -1
			If oItems(iIdx(j)) < oItems(iIdx(j-1)) Then
				temp = iIdx(j)
				iIdx(j) = iIdx(j-1)
				iIdx(j-1) = temp
				bChanged = True
			End If
		Next
		If Not bChanged Then Exit For
	Next
End Sub

'Listing 550. Global variables defined in the Inspector module
Private oDlg 			'Displayed dialog
Private oProgress 		'Progress control model
Private oTextEdit 		'Text edit control model that displays information
Private oObjects(100) 	'Objects to inspect
Private Titles(100) 	'Titles
Private iDepth% 		'Which object / title to use

'Listing 551. Create the dialog, the controls, and the listeners
Sub Inspect(Optional oInsObj, Optional title$)
	Dim oDlgModel 				'The dialog's model
	Dim oModel 					'Model for a control
	Dim oListener 				'A created listener object
	Dim oControl 				'References a control
	Dim iTabIndex As Integer 	'The current tab index while creating controls
	Dim iDlgHeight As Long 		'The dialog's height
	Dim iDlgWidth As Long 		'The dialog's width
	Dim sTitle$

	iDepth = LBound(oObjects)
	sTitle = ""
	If Not IsMissing(title) Then sTitle = title

	REM If no object is passed in, then use ThisComponent.
	If IsMissing(oInsObj) Then
		oObjects(iDepth) = ThisComponent
		If IsMissing(title) Then sTitle = "ThisComponent"
	Else
		oObjects(iDepth) = oInsObj
		If IsMissing(title) Then sTitle = "Obj"
	End If
	Titles(iDepth) = sTitle

	iDlgHeight = 300
	iDlgWidth = 350

	REM Create the dialog's model
	oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	setProperties(oDlgModel, Array("PositionX", 50, "PositionY", 50,_
		"Width", iDlgWidth, "Height", iDlgHeight, "Title", sTitle))

	createInsertControl(oDlgModel, iTabIndex, "PropButton",_
		"com.sun.star.awt.UnoControlRadioButtonModel",_
		Array("PositionX", 10, "PositionY", 10, "Width", 50, "Height", 15,_
			"Label", "Properties"))

	createInsertControl(oDlgModel, iTabIndex, "MethodButton",_
		"com.sun.star.awt.UnoControlRadioButtonModel",_
		Array("PositionX", 65, "PositionY", 10, "Width", 50, "Height", 15,_
			"Label", "Methods"))

	createInsertControl(oDlgModel, iTabIndex, "ServiceButton",_
		"com.sun.star.awt.UnoControlRadioButtonModel",_
		Array("PositionX", 120, "PositionY", 10, "Width", 50, "Height", 15,_
			"Label", "Services"))
	
	createInsertControl(oDlgModel, iTabIndex, "ObjectButton",_
		"com.sun.star.awt.UnoControlRadioButtonModel",_
		Array("PositionX", 175, "PositionY", 10, "Width", 50, "Height", 15,_
			"Label", "Object"))

	createInsertControl(oDlgModel, iTabIndex, "EditControl",_
		"com.sun.star.awt.UnoControlEditModel",_
		Array("PositionX", 10, "PositionY", 25, "Width", iDlgWidth - 20,_
			"Height", (iDlgHeight - 75), "HScroll", True, "VScroll", True,_
			"MultiLine", True, "HardLineBreaks", True))

	REM Store the edit control's model in a global variable.
	oTextEdit = oDlgModel.getByName("EditControl")

	createInsertControl(oDlgModel, iTabIndex, "Progress",_
		"com.sun.star.awt.UnoControlProgressBarModel",_
		Array("PositionX", 10, "PositionY", (iDlgHeight - 45),_
			"Width", iDlgWidth - 20, "Height", 15, "ProgressValueMin", 0,_
			"ProgressValueMax", 100))
	
	REM Store a reference to the progress bar
	oProgress = oDlgModel.getByName("Progress")

	REM Notice that I set the type to OK so I do not require an action
	REM listener to close the dialog.
	createInsertControl(oDlgModel, iTabIndex, "OKButton",_
		"com.sun.star.awt.UnoControlButtonModel",_
		Array("PositionX", Clng(25), "PositionY", iDlgHeight-20,_
			"Width", 50, "Height", 15, "Label", "Close",_
			"PushButtonType", com.sun.star.awt.PushButtonType.OK))

	createInsertControl(oDlgModel, iTabIndex, "InspectButton",_
		"com.sun.star.awt.UnoControlButtonModel",_
		Array("PositionX", Clng(iDlgWidth / 2 - 50), "PositionY", iDlgHeight-20,_
		"Width", 50, "Height", 15, "Label", "Inspect Selected",_
		"PushButtonType", com.sun.star.awt.PushButtonType.STANDARD))

	createInsertControl(oDlgModel, iTabIndex, "BackButton",_
		"com.sun.star.awt.UnoControlButtonModel",_
		Array("PositionX", Clng(iDlgWidth / 2 + 50), "PositionY", iDlgHeight-20,_
		"Width", 50, "Height", 15, "Label", "Inspect Previous",_
		"Enabled", False, _
		"PushButtonType", com.sun.star.awt.PushButtonType.STANDARD))

	REM Create the dialog and set the model
	oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oDlg.setModel(oDlgModel)

	REM the item listener for all of the radio buttons.
	oListener = CreateUnoListener("radio_", "com.sun.star.awt.XItemListener")
	oControl = oDlg.getControl("PropButton")
	ocontrol.addItemListener(oListener)
	oControl = oDlg.getControl("MethodButton")
	ocontrol.addItemListener(oListener)
	oControl = oDlg.getControl("ServiceButton")
	ocontrol.addItemListener(oListener)
	oControl = oDlg.getControl("ObjectButton")
	ocontrol.addItemListener(oListener)
	oControl.getModel().State = 1

	oListener = CreateUnoListener("ins_", "com.sun.star.awt.XActionListener")
	oControl = oDlg.getControl("InspectButton")
	ocontrol.addActionListener(oListener)

	oListener = CreateUnoListener("back_", "com.sun.star.awt.XActionListener")
	oControl = oDlg.getControl("BackButton")
	ocontrol.addActionListener(oListener)

	REM Now, set the dialog to contain the standard dialog information
	DisplayNewObject()

	REM Create a window and then tell the dialog to use the created window.
	Dim oWindow
	oWindow = CreateUnoService("com.sun.star.awt.Toolkit")
	oDlg.createPeer(oWindow, null)

	REM Finally, execute the dialog.
	oDlg.execute()
End Sub

'Listing 552. Create a control model and insert it into the dialog's model
'*************************************************************************
'** Create a control of type sType with name sName, and insert it into
'** oDlgModel.
'**
'** oDlgModel - Dialog model into which the control will be placed.
'** index - Tab index, which is incremented each time.
'** sName - Control name.
'** sType - Full service name for the control model.
'** props - Arry of property name / property value pairs.
'*************************************************************************
Sub createInsertControl(oDlgModel, index%, sName$, sType$, props())
	Dim oModel

	oModel = oDlgModel.createInstance(sType$)
	setProperties(oModel, props())
	setProperties(oModel, Array("Name", sName$, "TabIndex", index%))
	oDlgModel.insertByName(sName$, oModel)

	REM This changes the value because it is not passed by value.
	index% = index% + 1
End Sub

'Listing 553. Set many properties at one time
REM Generically set properties based on an array of name/value pairs.
Sub setProperties(oModel, props())
	Dim i As Integer
	For i=LBound(props()) To UBound(props()) Step 2
		oModel.setPropertyValue(props(i), props(i+1))
	Next
End sub

'Listing 554. The event handler is very simple; it calls DisplayNewObject()
REM This method is to support the com.sun.star.awt.XItemListener interface
REM for the radio button listener. This method is called when a radio button
REM is selected.
Sub radio_itemStateChanged(oItemEvent)
	DisplayNewObject()
End Sub

Sub DisplayNewObject()
	REM Reset the progress bar!
	oProgress.ProgressValue = 0
	oTextEdit.Text = ""

	On Local Error GoTo IgnoreError:
	If IsNull(oObjects(iDepth)) OR IsEmpty(oObjects(iDepth)) Then
		oTextEdit.Text = ObjInfoString(oObjects(iDepth))
	ElseIf oDlg.getModel().getByName("PropButton").State = 1 Then
		processStateChange(oObjects(iDepth), "p")
	ElseIf oDlg.getModel().getByName("MethodButton").State = 1 Then
		processStateChange(oObjects(iDepth), "m")
	ElseIf oDlg.getModel().getByName("ServiceButton").State = 1 Then
		processStateChange(oObjects(iDepth), "s")
	Else
		oTextEdit.Text = ObjInfoString(oObjects(iDepth))
	End If
	oProgress.ProgressValue = 100

	IgnoreError:
End Sub

'Listing 555. Inspect a new item
Sub Ins_actionPerformed(oActionEvent)
	Dim oControl
	Dim oSel
	Dim sText$
	Dim v
	Dim oOldObj
	Dim sOldTitle$
	Dim sNewTitle$

	If iDepth = UBound(oObjects) Then
		Print "Sorry, already nested too deeply"
		Exit Sub
	End If

	oControl = oDlg.getControl("EditControl")
	sText = oControl.getSelectedText()
	sOldTitle = oDlg.Title
	sNewTitle = sOldTitle & "/" & sText

	If IsNumeric(sText) Then
		' oDlg.getModel().getByName("ObjButton").State = 1
		If NumArrayDimensions(oObjects(iDepth)) <> 1 Then
			Print "You can only select a number with a one dimensional array."
		Else
			oOldObj = oObjects(iDepth)
			iDepth = iDepth + 1
			oDlg.getControl("BackButton").GetModel().Enabled = True
			oObjects(iDepth) = oOldObj(sText)
			oDlg.Title = sNewTitle
			Titles(iDepth) = sNewTitle
			DisplayNewObject()
		End If
	Else
		v = RunFromLib(GlobalScope.BasicLibraries, _
					"xyzzylib", "TestMod", oObjects(iDepth), sText, False, False)
		If IsNull(v) OR IsEmpty(v) Then
			'
		Else
			iDepth = iDepth + 1
			oDlg.getControl("BackButton").GetModel().Enabled = True
			oObjects(iDepth) = v
			oDlg.Title = sNewTitle
			Titles(iDepth) = sNewTitle
			DisplayNewObject()
		End If
	End If
	'oSel = oControl.getSelection ()
	'Print "Min: " & oSel.Min & " Max: " & oSel.Max
	'sText = oControl.getText()
	'Inspect oControl
End Sub

'Listing 556. Inspect a previous item
Sub back_actionPerformed(oActionEvent)
	If iDepth <= LBound(oObjects) Then
		Exit Sub
	End If
	iDepth = iDepth - 1
	If iDepth = LBound(oObjects) Then
		oDlg.getControl("BackButton").GetModel().Enabled = False
	End If

	oDlg.Title = Titles(iDepth)
	DisplayNewObject()
End Sub

'Listing 557. Build the text string that is displayed about the object
'*************************************************************************
'** Set the text. For a service, show the interfaces and services in
'** Separate sections so that they are easier to find.
'*************************************************************************
Sub processStateChange(oInsObj, sPropType$)
	Dim oItems()
	BuildItemArray(oInsObj, sPropType$, oItems())
	If sPropType$ = "s" Then
		Dim s As String
		On Local Error Resume Next
		s = "** INTERFACES **" & CHR$(10) & Join(oItems, CHR$(10))
		s = s & CHR$(10) & CHR$(10) & "** SERVICES **" & CHR$(10) & _
			Join(oInsObj.getSupportedServiceNames(), CHR$(10))
		oTextEdit.Text = s
	Else
		oTextEdit.Text = Join(oItems, CHR$(10))
	End If
End Sub

'Listing 558. Build an array of properties, methods, or services
'*************************************************************************
'** This routine parses the strings returned from the dbg_Methods,
'** dbg_Properties, and Dbg_SupportedInterfaces calls. The interesting
'** data starts after the first colon.
'**
'** Because of this, all data before the first colon is discarded
'** and then the string is separated into pieces based on the
'** separator string that is passed in.
'**
'** All instances of the string Sbx are removed. If this string
'** is valid and exists in a method name, it will still be removed so perhaps
'** it is not the safest thing to do, but I have never seen this case and it
'** makes the output easier to read.
'**
'** the oItems() contains all of the parsed sections on output.
'*************************************************************************
Sub BuildItemArray(oInsObj, sType$, oItems())
	On Error Goto BadErrorHere
	Dim s As String 			'Primary list to parse
	Dim sSep As String 			'Separates the items in the string
	Dim iCount% '
	Dim iPos% '
	Dim sNew$ '
	Dim i% '
	Dim j% '
	Dim sFront() As String 		'When each piece is parsed, this is the front
	Dim sMid() As String 		'When each piece is parsed, this is the middle
	Dim iIdx() As Integer 		'Used to sort parallel arrays.
	Dim nFrontMax As Integer 	'Maximum length of front section
	Dim sArraySep$ 				: sArraySep = ", "
	Dim iArrayMaxDepth% 		: iArrayMaxDepth = 10

	nFrontMax = 0
	
	REM First, see what should be inspected.
	s = ""
	On Local Error Resume Next
	If sType$ = "s" Then
		REM Dbg_SupportedInterfaces returns interfaces and
		REM getSupportedServiceNames() returns services.
		s = oInsObj.Dbg_SupportedInterfaces
		's = s & Join(oInsObj.getSupportedServiceNames(), CHR$(10))
		sSep = CHR$(10)
	ElseIf sType$ = "m" Then
		s = oInsObj.DBG_Methods
		sSep = ";"
	ElseIf sType$ = "p" Then
		s = oInsObj.DBG_Properties
		sSep = ";"
	Else
		s = ""
		sSep = ""
	End If

	REM The dbg_ variables have some introductory information that
	REM we do not want so remove it.
	REM We only care about what is after the colon.
	iPos = InStr(1, s, ":") + 1
	If iPos > 0 Then s = TrimWhite(Right(s, Len(s) - iPos))
	
	REM All data types are prefixed with the text Sbx.
	REM Remove all of the "SbX" charcters
	s = Join(Split(s, "Sbx"), "")
	
	REM If the separator is NOT CHR$(10), then remove
	REM all instances of CHR$(10)
	If ASC(sSep) <> 10 Then s = Join(Split(s, CHR$(10)), "")

	REM split on the separator character and update the progress bar.
	oItems() = Split(s, sSep)
	oProgress.ProgressValue = 20
	
	Rem Create arrays to hold the different portions of the text.
	Rem the string usually contains text similar to "SbxString getName()"
	Rem sFront() holds the data type if it exists and "" if it does not.
	Rem sMid() holds the rest
	ReDim sFront(UBound(oItems)) As String
	ReDim sMid(UBound(oItems)) As String
	ReDim iIdx(UBound(oItems)) As Integer

	REM Initialize the index array and remove leading and trailing
	REM spaces from each string.
	For i=LBound(oItems()) To UBound(oItems())
		oItems(i) = Trim(oItems(i))
		iIdx(i) = i
		j = InStr(1, oItems(i), " ")
		If (j > 0) Then
			REM If the string contains more than one word, the first word is stored
			REM in sFront() and the rest of the string is stored in sMid().
			sFront(i) = Mid$(oItems(i), 1, j)
			sMid(i) = Mid$(oItems(i), j+1)
			If j > nFrontMax Then nFrontMax = j
		Else
			REM If the string contains only one word, sFront() is empty
			REM and the string is stored in sMid().
			sFront(i) = ""
			sMid(i) = oItems(i)
		End If
	Next

	oProgress.ProgressValue = 40
	Rem Sort the primary names. The array is left unchanged, but the
	Rem iIdx() array contains index values that allow a sorted traversal
	SortMyArray(sMid(), iIdx())
	oProgress.ProgressValue = 50

	REM Dealing with properties so attempt to find the value
	REM of each property.
	If sType$ = "p" Then
		Dim oPropInfo 'PropertySetInfo object
		Dim oProps 'Array of properties
		Dim oProp 'com.sun.star.beans.Property
		Dim v 'Value of a single property
		Dim bHasPI As Boolean
		Dim bUsePI As Boolean
		Dim bConvertReturns As Boolean
		
		bConvertReturns = True
		bHasPI = False
		bUsePI = False
		On Error Goto NoPropertySetInfo
		oPropInfo = oInsObj.getPropertySetInfo()
		bHasPI = True
		NoPropertySetInfo:

		On Error Goto BadErrorHere:
		For i=LBound(sMid()) To UBound(sMid())
			If bHasPI Then
				bUsePI = oPropInfo.hasPropertyByName(sMid(i))
			End If
			
			If bUsePI Then
				v = oInsObj.getPropertyValue(sMid(i))
			Else
				v = RunFromLib(GlobalScope.BasicLibraries, _
						"xyzzylib", "TestMod", oInsObj, sMid(i), False, False)
			End If
			sMid(i) = sMid(i) & " = " & _
					ObjToString(v, sArraySep, iArrayMaxDepth, bConvertReturns)
		Next
	End If
	oProgress.ProgressValue = 60

	nFrontMax = nFrontMax + 1
	iCount = LBound(oItems())
	REM Now build the array of the items in sorted order
	REM Sometimes, a service is listed more than once.
	REM this routine removes multiple instances of the same service.
	For i = LBound(oItems()) To UBound(oItems())
		sNew = sFront(iIdx(i)) & " " & sMid(iIdx(i))
		'Uncomment these lines if you want to add uniform space
		'between front and mid. This is only useful if the font is fixed in width.
		'sNew = sFront(iIdx(i))
		'sNew = sNew & Space(nFrontMax - Len(sNew)) & sMid(iIdx(i))
		If i = LBound(oItems()) Then
			oItems(iCount) = sNew
		ElseIf oItems(iCount) <> sNew Then
			iCount = iCount + 1
			oItems(iCount) = sNew
		End If
	Next
	oProgress.ProgressValue = 75
	ReDim Preserve oItems(iCount)
	Exit Sub
BadErrorHere:
	MsgBox "Error " & err & ": " & error$ + chr(13) + "In line : " + Erl
End Sub

'Listing 559. Create a library, module, and function, then call it
'*************************************************************************
'** Create a module that contains a function that returns the value from a
'** property or makes a method call on an object.
'**
'** Not all properties are available from the PropertySetInfo object, and
'** I am not aware of any other way to make a method call.
'**
'** oLibs - Library container to use
'** sLName - Library name
'** sMName - Module name
'** oObj - Object of interest
'** sCall - Method or property name to call
'** bClean - If true, remove the module. Library is only removed if created.
'** bIsMthd- If true, create as a method rather than a property.
'** x - If exists, then used as parameter for the method call.
'** y - If exists, then used as parameter for the method call.
'**
'*************************************************************************
Function RunFromLib(oLibs, sLName$, sMName$, oObj, sCall$, _
					bClean As Boolean, bIsMthd As Boolean, _
					Optional x, optional y)
	Dim oLib 	'The library to use to run the new function.
	Dim s$ 		'Generic string variable.
	Dim bAddedLib As Boolean

	REM If the library does not exist, then create it.
	bAddedLib = False
	If NOT oLibs.hasByName(sLName) Then
		bAddedLib = True
		oLibs.createLibrary(sLName)
	End If
	oLibs.loadLibrary(sLName)
	oLib = oLibs.getByName(sLName)

	If oLib.hasByName(sMName) Then
		oLib.removeByName(sMName)
	End If

	s = "Option Explicit" & CHR$(10) & _
		"Function MyInsertedFunc(ByRef oObj"

	If NOT IsMissing(x) Then s = s & ", x"
	If NOT IsMissing(y) Then s = s & ", y"
	s = s & ")" & CHR$(10) & "On Local Error Resume Next" & CHR$(10)
	s = s & "MyInsertedFunc = oObj." & sCall

	If bIsMthd Then
		s = s & "("
		If NOT IsMissing(x) Then s = s & "x"
		If NOT IsMissing(y) Then s = s & ", y"
		s = s & ")"
	End If

	s = s & CHR$(10) & "End Function"

	oLib.insertByName(sMName, s)

	If IsMissing(x) Then
		RunFromLib = MyInsertedFunc(oObj)
	ElseIf IsMissing(y) Then
		RunFromLib = MyInsertedFunc(oObj, x)
	Else
		RunFromLib = MyInsertedFunc(oObj, x, y)
	End If

	If bClean Then
		oLib.removeByName(sMName)
		If bAddedLib Then
			oLibs.removeLibrary(sLName)
		End If
	End If
End Function

