Public Class clsCEC
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbouiCOM.Item
    Private oInvoice As SAPbobsCOM.Documents
    Private ofolder As SAPbouiCOM.Folder
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub AddControl(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        oApplication.Utilities.AddControls(aform, "FldAt", "39", SAPbouiCOM.BoFormItemTypes.it_FOLDER, "RIGHT", 0, 0, "39", "CEC Attributes")
        Dim oldItem As SAPbouiCOM.Item
        oApplication.Utilities.AddControls(aform, "grdAt", "98", SAPbouiCOM.BoFormItemTypes.it_GRID, "DOWN", 7, 7, , , 200, , 80)
        oItem = aform.Items.Item("grdAt")
        oldItem = aform.Items.Item("98")
        oItem.Top = oldItem.Top ' + 20
        oItem.Width = oldItem.Width
        oItem.Height = oldItem.Height
        oApplication.Utilities.AddControls(aform, "btnAdd1", "83", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 7, 7, , "Add Row")
        oApplication.Utilities.AddControls(aform, "btnDel1", "btnAdd1", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 7, 7, "btnAdd1", "Delete Row")
        aform.DataSources.DataTables.Add("dtAtt")
        oItem = aform.Items.Item("FldAt")
        oItem.AffectsFormMode = False
        ofolder = oItem.Specific
        ofolder.GroupWith("39")
        ofolder.ValOn = "X"
        ofolder.ValOff = ""
        LoadGridValue(oForm)
        aform.Freeze(False)
    End Sub

    Private Function AddtoUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUsertable As SAPbobsCOM.UserTable
        Dim strsql, code, Name, strCode, strItemcode, strSerialNo, strCardCode, strMfrSerialNo As String
        Dim aGrid As SAPbouiCOM.Grid
        Try
            aform.Freeze(True)
            aGrid = aform.Items.Item("grdAt").Specific
            oUsertable = oApplication.Company.UserTables.Item("Z_OCECATT")
            strItemcode = oApplication.Utilities.getEdittextvalue(aform, "45")
            strSerialNo = oApplication.Utilities.getEdittextvalue(aform, "44")
            strCardCode = oApplication.Utilities.getEdittextvalue(aform, "48")
            strMfrSerialNo = oApplication.Utilities.getEdittextvalue(aform, "43")
            For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                If aGrid.DataTable.GetValue("U_Z_Code", intRow) <> "" Then
                    If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        strCode = oApplication.Utilities.getMaxCode("@Z_OCECATT", "Code")
                        oUsertable.Code = strCode
                        oUsertable.Name = strCode & "_N"
                        oUsertable.UserFields.Fields.Item("U_Z_ItemCode").Value = strItemcode
                        oUsertable.UserFields.Fields.Item("U_Z_SerialNo").Value = strSerialNo
                        oUsertable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                        oUsertable.UserFields.Fields.Item("U_Z_manufSN").Value = strMfrSerialNo
                        oUsertable.UserFields.Fields.Item("U_Z_Code").Value = aGrid.DataTable.GetValue("U_Z_Code", intRow)
                        oUsertable.UserFields.Fields.Item("U_Z_Desc").Value = aGrid.DataTable.GetValue("U_Z_Desc", intRow)
                        If oUsertable.Add <> 0 Then
                        Else
                            aGrid.DataTable.SetValue("Code", intRow, strCode)
                        End If
                    Else
                        If oUsertable.GetByKey(aGrid.DataTable.GetValue("Code", intRow)) = False Then
                            strCode = oApplication.Utilities.getMaxCode("@Z_OCECATT", "Code")
                            oUsertable.Code = strCode
                            oUsertable.Name = strCode & "_N"
                            oUsertable.UserFields.Fields.Item("U_Z_ItemCode").Value = strItemcode
                            oUsertable.UserFields.Fields.Item("U_Z_SerialNo").Value = strSerialNo
                            oUsertable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                            oUsertable.UserFields.Fields.Item("U_Z_manufSN").Value = strMfrSerialNo
                            oUsertable.UserFields.Fields.Item("U_Z_Code").Value = aGrid.DataTable.GetValue("U_Z_Code", intRow)
                            oUsertable.UserFields.Fields.Item("U_Z_Desc").Value = aGrid.DataTable.GetValue("U_Z_Desc", intRow)
                            If oUsertable.Add <> 0 Then
                            Else
                                aGrid.DataTable.SetValue("Code", intRow, strCode)
                            End If
                        Else
                            strCode = aGrid.DataTable.GetValue("Code", intRow)
                            oUsertable.Code = strCode
                            oUsertable.Name = strCode
                            oUsertable.UserFields.Fields.Item("U_Z_ItemCode").Value = strItemcode
                            oUsertable.UserFields.Fields.Item("U_Z_SerialNo").Value = strSerialNo
                            oUsertable.UserFields.Fields.Item("U_Z_CardCode").Value = strCardCode
                            oUsertable.UserFields.Fields.Item("U_Z_manufSN").Value = strMfrSerialNo
                            oUsertable.UserFields.Fields.Item("U_Z_Code").Value = aGrid.DataTable.GetValue("U_Z_Code", intRow)
                            oUsertable.UserFields.Fields.Item("U_Z_Desc").Value = aGrid.DataTable.GetValue("U_Z_Desc", intRow)
                            If oUsertable.Update <> 0 Then
                            Else
                                aGrid.DataTable.SetValue("Code", intRow, strCode)
                            End If
                        End If
                    End If
                   
                End If
            Next
            ' LoadGridValue(aform)
            aform.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End Try
    End Function
    Private Sub LoadGridValue(ByVal aform As SAPbouiCOM.Form, Optional ByVal aChoice As String = "Normal")
        Dim strsql As String
        Dim oRec As SAPbobsCOM.Recordset
        Try
            aform.Freeze(True)
            If aform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                strsql = "SELECT T0.[insID] FROM OINS T0 WHERE 1=2"
            End If
            Dim intContid As Integer
            If aChoice <> "Duplicate" Then
                strsql = "SELECT T0.[insID] FROM OINS T0 WHERE T0.[manufSN]='" & oApplication.Utilities.getEdittextvalue(aform, "43") & "' and  T0.[internalSN] ='" & oApplication.Utilities.getEdittextvalue(aform, "44") & "' and  T0.[itemCode]  ='" & oApplication.Utilities.getEdittextvalue(aform, "45") & "' and  T0.[customer] ='" & oApplication.Utilities.getEdittextvalue(aform, "48") & "'"
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(strsql)

                If oRec.RecordCount > 0 Then
                    intContid = oRec.Fields.Item(0).Value
                Else
                    intContid = 0
                End If
            Else
                intContid = 0
            End If
          
            '    = oRec.Fields.Item(0).Value
            oGrid = aform.Items.Item("grdAt").Specific
            oGrid.DataTable = aform.DataSources.DataTables.Item("dtAtt")
            oGrid.DataTable.ExecuteQuery("Select * from [@Z_OCECATT] where U_Z_DocEntry='" & intContid & "'")
            ' oGrid.DataTable.ExecuteQuery("Select * from [@Z_OCECATT] T0 where T0.[U_Z_SerialNo] ='" & oApplication.Utilities.getEdittextvalue(aform, "44") & "' and  T0.[U_Z_ItemCode]  ='" & oApplication.Utilities.getEdittextvalue(aform, "45") & "' and  T0.[U_Z_CardCode] ='" & oApplication.Utilities.getEdittextvalue(aform, "48") & "'")

            oGrid.Columns.Item("Code").Visible = False
            oGrid.Columns.Item("Name").Visible = False
            oGrid.Columns.Item("U_Z_DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocNum").Visible = False

            oGrid.Columns.Item("U_Z_Code").TitleObject.Caption = "Code"
            oGrid.Columns.Item("U_Z_Code").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombobox = oGrid.Columns.Item("U_Z_Code")
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select * from [@Z_OATT] order by Code")
            For intRow As Integer = 0 To oTest.RecordCount - 1
                oCombobox.ValidValues.Add(oTest.Fields.Item("Code").Value, oTest.Fields.Item("Name").Value)
                oTest.MoveNext()
            Next
            oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Value
            oGrid.Columns.Item("U_Z_Desc").TitleObject.Caption = "Name"
            oGrid.Columns.Item("U_Z_Desc").Editable = False
            oGrid.Columns.Item("U_Z_ItemCode").Visible = False
            oGrid.Columns.Item("U_Z_SerialNo").Visible = False
            oGrid.Columns.Item("U_Z_CardCode").Visible = False
            oGrid.Columns.Item("U_Z_manufSN").Visible = False
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            For intLoop As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intLoop, intLoop + 1)
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CEC Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    CommitTrans("Cancel")
                                End If
                                If pVal.ItemUID = "3" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControl(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "grdAt" And pVal.ColUID = "U_Z_Code" Then
                                    oGrid = oForm.Items.Item("grdAt").Specific
                                    oCombobox = oGrid.Columns.Item("U_Z_Code")
                                    oGrid.DataTable.SetValue("U_Z_Desc", pVal.Row, oCombobox.GetSelectedValue(pVal.Row).Description)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "FldAt" Then
                                    oForm.PaneLevel = 7
                                End If
                                If pVal.ItemUID = "btnAdd1" Then
                                    oGrid = oForm.Items.Item("grdAt").Specific
                                    AddEmptyRow(oGrid)
                                End If
                                If pVal.ItemUID = "btnDel1" Then
                                    oGrid = oForm.Items.Item("grdAt").Specific
                                    DeleteRow(oForm, oGrid)
                                End If

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        If aGrid.DataTable.Rows.Count - 1 <= 0 Then
            aGrid.DataTable.Rows.Add()
            '  aGrid.Columns.Item("U_Z_Code").Click(aGrid.DataTable.Rows.Count - 1, False)
        End If
        If aGrid.DataTable.GetValue("U_Z_Code", aGrid.DataTable.Rows.Count - 1) <> "" Then
            aGrid.DataTable.Rows.Add()
            ' aGrid.Columns.Item("U_Z_Code").Click(aGrid.DataTable.Rows.Count - 1, False)
        End If
        For intLoop As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intLoop, intLoop + 1)
        Next
    End Sub
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form, ByVal aGrid As SAPbouiCOM.Grid)
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            If aGrid.Rows.IsSelected(intRow) Then
                Dim oTest As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("Update [@Z_OCECATT] set Name =Name + '_XD' where Code='" & aGrid.DataTable.GetValue("Code", intRow) & "'")
                aGrid.DataTable.Rows.Remove(intRow)
                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
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
            oTest.DoQuery("Delete from [@Z_OCECATT] where Name like '%_XD'")
            oTest.DoQuery("Update [@Z_OCECATT] set Name=Code where Name like '%_N'")
        Else
            oTest.DoQuery("Delete from  [@Z_OCECATT]  where Name like '%_N'")
            oTest.DoQuery("Update [@Z_OCECATT] set Name=Code where Name like '%_XD'")
        End If
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strCode, strCode1 As String
        oGrid = aForm.Items.Item("grdAt").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCombobox = oGrid.Columns.Item("U_Z_Code")
            strCode = oCombobox.GetSelectedValue(intRow).Value
            If strCode <> "" Then
                For intloop As Integer = intRow + 1 To oGrid.DataTable.Rows.Count - 1
                    strCode1 = oCombobox.GetSelectedValue(intloop).Value
                    If strCode1 <> "" And strCode.ToUpper = strCode1.ToUpper Then
                        oApplication.Utilities.Message("This attribute already exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item("U_Z_Code").Click(intloop)
                        Return False
                    End If

                Next
            End If
        Next
        Return True


    End Function

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "12"
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
                        DeleteRow(oForm, oGrid)
                        BubbleEvent = False
                        Exit Sub
                    End If

                Case "1287"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("1").Specific
                    If pVal.BeforeAction = False Then
                        LoadGridValue(oForm, "Duplicate")
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then

                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                '  If oForm.TypeEx = frm_CEC And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                LoadGridValue(oForm)
                'End If
            ElseIf BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                Dim st As SAPbobsCOM.CustomerEquipmentCards
                Dim str As String
                st = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
                If st.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) = True Then
                    Dim ost As SAPbobsCOM.Recordset
                    ost = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If AddtoUDT(oForm) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    str = "Update [@Z_OCECATT] set U_Z_DocEntry='" & st.EquipmentCardNum & "' where U_Z_manufSN='" & st.ManufacturerSerialNum & "' and  U_Z_ItemCode='" & st.ItemCode & "' and U_Z_SerialNo='" & st.InternalSerialNum & "' and U_Z_CardCode='" & st.CustomerCode & "'"
                    ost.DoQuery(str)
                End If
                CommitTrans("Add")
                LoadGridValue(oForm)
            ElseIf BusinessObjectInfo.BeforeAction = True And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                'If AddtoUDT(oForm) = False Then
                '    BubbleEvent = False
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
