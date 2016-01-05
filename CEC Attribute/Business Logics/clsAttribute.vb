Public Class clsAttribute
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_CECAttribute, frm_CECAttribute)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oGrid = oForm.Items.Item("1").Specific
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        'AddChooseFromList(oForm)
        oForm.Freeze(True)
        FormatGrid(oGrid)
        oForm.Freeze(False)
    End Sub

    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strCode, strCode1, strName, strEname As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strCode = aGrid.DataTable.GetValue("Code", intRow)
            If strCode <> "" Then
                strName = aGrid.DataTable.GetValue("Name", intRow)
                If aGrid.DataTable.GetValue("Name", intRow) = "" Then
                    oApplication.Utilities.Message("Name is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aGrid.Columns.Item("Name").Click(intRow)
                    Return False
                End If

                For intLoop As Integer = intRow + 1 To aGrid.DataTable.Rows.Count - 1
                    strCode1 = aGrid.DataTable.GetValue("Code", intLoop)
                    If strCode1 <> "" Then
                        strEname = aGrid.DataTable.GetValue("Name", intLoop)
                        If strCode.ToUpper() = strCode1.ToUpper() Then
                            oApplication.Utilities.Message("This Code already exists : " & strCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aGrid.Columns.Item("Code").Click(intLoop)
                            Return False
                        End If

                        If strName.ToUpper() = strEname.ToUpper() Then
                            oApplication.Utilities.Message("This Name already exists : " & strCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            aGrid.Columns.Item("Code").Click(intLoop)
                            Return False
                        End If
                    End If
                Next
            End If

        Next
        Return True
    End Function
    Private Sub AddtoUDT(ByVal aGrid As SAPbouiCOM.Grid)
        Dim oUsertable As SAPbobsCOM.UserTable
        Dim strsql, code, Name, strCode As String
        oUsertable = oApplication.Company.UserTables.Item("Z_OATT")
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            If aGrid.DataTable.GetValue("Code", intRow) <> "" Then


                If oUsertable.GetByKey(aGrid.DataTable.GetValue("Code", intRow)) = False Then
                    oUsertable.Code = aGrid.DataTable.GetValue("Code", intRow)
                    oUsertable.Name = aGrid.DataTable.GetValue("Name", intRow)
                        If oUsertable.Add <> 0 Then
                    End If
                Else
                    oUsertable.Code = aGrid.DataTable.GetValue("Code", intRow)
                    oUsertable.Name = aGrid.DataTable.GetValue("Name", intRow)
                    If oUsertable.Update <> 0 Then
                    End If
                End If
            End If
        Next
        CommitTrans("Add")
    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        oGrid.DataTable.ExecuteQuery("select *  from [@Z_OATT] order by Code")
        oGrid.Columns.Item("Code").TitleObject.Caption = "Code"
        oGrid.Columns.Item("Name").TitleObject.Caption = "Name"
        oGrid.AutoResizeColumns()
        For intLoop As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intLoop, intLoop + 1)
        Next
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        If aGrid.DataTable.GetValue("Code", aGrid.DataTable.Rows.Count - 1) <> "" Then
            aGrid.DataTable.Rows.Add()
            aGrid.Columns.Item(0).Click(aGrid.DataTable.Rows.Count - 1, False)
        End If
        For intLoop As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intLoop, intLoop + 1)
        Next
    End Sub
    Private Sub DeleteRow(ByVal aGrid As SAPbouiCOM.Grid)
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            If aGrid.Rows.IsSelected(intRow) Then
                Dim oTest As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("Select * from [@Z_OCECATT] where U_Z_Code='" & aGrid.DataTable.GetValue("Code", intRow) & "'")
                If oTest.RecordCount > 0 Then
                    oApplication.Utilities.Message("Attribute already mapped in Customer Equipement Card", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                oTest.DoQuery("Update [@Z_OATT] set Name =Name + '_XD' where Code='" & aGrid.DataTable.GetValue("Code", intRow) & "'")
                aGrid.DataTable.Rows.Remove(intRow)
                For intLoop As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                    aGrid.RowHeaders.SetText(intLoop, intLoop + 1)
                Next
                Exit Sub
            End If
        Next
    End Sub

    Private Sub CommitTrans(ByVal aChoice As String)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aChoice = "Add" Then
            oTest.DoQuery("Delete from [@Z_OATT] where Name like '%_XD'")
        Else
            oTest.DoQuery("Update [@Z_OATT] set Name=replace(Name,'_XD','')")
        End If
    End Sub


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CECAttribute Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    oForm.Freeze(True)
                                    oGrid = oForm.Items.Item("1").Specific
                                    If validation(oGrid) = False Then
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    oGrid = oForm.Items.Item("1").Specific
                                    AddtoUDT(oGrid)
                                    FormatGrid(oGrid)
                                    oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oForm.Freeze(False)
                                End If
                                If pVal.ItemUID = "2" Then
                                    CommitTrans("Cancel")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And pVal.ColUID = "Code" Then
                                    Dim agrid As SAPbouiCOM.Grid
                                    Dim oTest As SAPbobsCOM.Recordset
                                    agrid = oForm.Items.Item("1").Specific
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oTest.DoQuery("Select * from [@Z_OCECATT] where U_Z_Code='" & agrid.DataTable.GetValue("Code", pVal.Row) & "'")
                                    If oTest.RecordCount > 0 Then
                                        oApplication.Utilities.Message("Attribute already mapped in Customer Equipement Card", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                    Dim agrid As SAPbouiCOM.Grid
                                    Dim oTest As SAPbobsCOM.Recordset
                                    agrid = oForm.Items.Item("1").Specific
                                    If agrid.DataTable.GetValue("Name", pVal.Row) <> "" Then
                                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oTest.DoQuery("Select * from [@Z_OCECATT] where U_Z_Code='" & agrid.DataTable.GetValue("Code", pVal.Row) & "'")
                                        If oTest.RecordCount > 0 Then
                                            oApplication.Utilities.Message("Attribute already mapped in Customer Equipement Card", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    'If (oCFLEvento.BeforeAction = False) Then
                                    '    Dim oDataTable As SAPbouiCOM.DataTable
                                    '    oDataTable = oCFLEvento.SelectedObjects
                                    '    oGrid = oForm.Items.Item("1").Specific
                                    '    intChoice = 0
                                    '    oForm.Freeze(True)
                                    '    If ((pVal.ItemUID = "1" And (pVal.ColUID = "U_Z_GLACC" Or pVal.ColUID = "U_Z_Credit"))) Then
                                    '        val = oDataTable.GetValue("FormatCode", 0)
                                    '        oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                    '    End If
                                    '    oForm.Freeze(False)
                                    'End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If


        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_CECAttribute
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("1").Specific
                    If pVal.BeforeAction = False Then
                        AddEmptyRow(oGrid)
                    End If

                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("1").Specific
                    If pVal.BeforeAction = True Then
                        DeleteRow(oGrid)
                        BubbleEvent = False
                        Exit Sub
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
