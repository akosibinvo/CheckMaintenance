'                                  Bryan     ./
'                                  AppTech (o o)
'--------------------------------------oOOo-(_)-oOOo---------------------------------
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Module Maintenance_File

    Dim oQuery As String
    Dim oCompany As SAPbobsCOM.Company
    Dim oRec As SAPbobsCOM.Recordset

    Dim oColumns As SAPbouiCOM.Columns
    Dim oColumn As SAPbouiCOM.Column

    'Dim dg_Less As SAPbouiCOM.Grid
    Dim oGrid As SAPbouiCOM.Grid
    Dim dg_Others As SAPbouiCOM.Grid
    Dim txt_NumB As SAPbouiCOM.EditText
    Dim txt_RateB As SAPbouiCOM.EditText
    Dim txt_AppT As SAPbouiCOM.EditText
    Dim txt_AppN As SAPbouiCOM.EditText
    Dim txt_AppD As SAPbouiCOM.EditText
    Dim txt_T As SAPbouiCOM.EditText
    Dim txtRM As SAPbouiCOM.EditText
    Dim txt_Code As SAPbouiCOM.EditText
    Dim txt_oCount As SAPbouiCOM.EditText
    Dim dt_Start As SAPbouiCOM.EditText
    Dim dt_End As SAPbouiCOM.EditText
    Dim cmb_Format As SAPbouiCOM.ComboBox
    Dim cmb_LoadT As SAPbouiCOM.ComboBox
    Dim cmb_CopyTo As SAPbouiCOM.ComboBox
    Dim txt_VMW As SAPbouiCOM.EditText
    Dim txt_Charge As SAPbouiCOM.EditText
    Dim txt_Charg2 As SAPbouiCOM.EditText
    Dim txt_Peri As SAPbouiCOM.EditText
    Dim txt_GMA As SAPbouiCOM.EditText
    Dim oBP As SAPbouiCOM.EditText
    Dim oCombo As SAPbouiCOM.ComboBoxColumn
    Dim fldr_FCR As SAPbouiCOM.Folder
    Dim oEditText As SAPbouiCOM.EditTextColumn
    Private oPrgress As SAPbouiCOM.ProgressBar
    Private frmMaintenance As SAPbouiCOM.Form
    Private isLoading As Boolean = False

#Region "ItemEvent"

    Friend Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        frmMaintenance = SAP_APP.SAP_Form

        Try

            Select Case pVal.Action_Success
                Case False
                Case True

                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                            Select Case pVal.ItemUID
                                Case "cmb_CopyTo"
                                    cmb_CopyTo = frmMaintenance.Items.Item("cmb_CopyTo").Specific
                                    oItems()
                                    txtRM.Value = cmb_CopyTo.Selected.Value.ToString
                                Case "cmb_Format"
                                    cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
                                    oItems()
                                    If txt_AppT.Value = "2" Then
                                        If cmb_Format.Value = "2" Then
                                            cmb_LoadT.Select(3, SAPbouiCOM.BoSearchKey.psk_Index)
                                        End If
                                    End If
                                    frmMaintenance.Items.Item("cmb_LoadT").Enabled = True

                                Case "cmb_LoadT"

                                    cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
                                    cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific

                                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim oLoad = cmb_LoadT.Value.ToString.Substring(0, 3)
                                    Dim oQuery As String = "SELECT * FROM [" & oCompany.CompanyDB & "]..[@APP_SERVICETYPE] WHERE Code LIKE '" & oLoad & "%'"

                                    oCombo = dgCOLUMNS.Columns.Item("Service Type")

                                    frmMaintenance.Freeze(True)
                                    For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                                        oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    oRec.DoQuery(oQuery)

                                    If oRec.RecordCount > 0 Then
                                        While oRec.EoF = False
                                            oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(1).Value.ToString)
                                            oRec.MoveNext()
                                        End While
                                    End If

                                    frmMaintenance.Items.Item("dg_Less").Visible = True
                                    frmMaintenance.Items.Item("dg_Full").Visible = False

                                    GoTo Unfreeze

                                Case "dg_Less"
                                    oCombo = dgCOLUMNS.Columns.Item("Service Type")
                                    oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                                Case "dg_Others"
                                    frmMaintenance.Freeze(True)
                                    Dim i As Integer = dgCOLUMNSOTHER.Rows.Count - 1
                                    oEditText = dgCOLUMNSOTHER.Columns.Item("Description")
                                    oCombo = dgCOLUMNSOTHER.Columns.Item("ItemCode")
                                    Dim oDesc As String = oCombo.GetSelectedValue(i).Description

                                    oEditText.SetText(i, oDesc)
                                    Try
                                        oEditText = dgCOLUMNSOTHER.Columns.Item("AccountCode")
                                        oRec.DoQuery("SELECT FormatCode FROM [" & oCompany.CompanyDB & "]..vintel_oitm WHERE ItemCode ='" & oCombo.GetSelectedValue(i).Value & "' ")
                                        If oRec.RecordCount > 0 Then
                                            oRec.MoveLast()
                                            'MsgBox(oRec.Fields.Item(0).Value.ToString)
                                            oEditText.SetText(i, oRec.Fields.Item(0).Value.ToString)
                                        End If

                                    Catch ex As Exception
                                        SAP_APP.SetMessage(ex)
                                    End Try
                                    GoTo Unfreeze
                            End Select
                        Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                            Select Case pVal.ItemUID
                                'For Binding
                                Case "tLoading"
                                    frmMaintenance.Freeze(True)

                                    frmMaintenance.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                    'frmMaintenance.Items.Item("Folder").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                    frmMaintenance.Items.Item("dt_Start").Click(SAPbouiCOM.BoCellClickType.ct_Double)

                                    Dim txt_C As SAPbouiCOM.EditText = frmMaintenance.Items.Item("txt_C").Specific
                                    Dim bt_Search As SAPbouiCOM.Button = frmMaintenance.Items.Item("bt_Search").Specific
                                    Dim txtRM As SAPbouiCOM.EditText = frmMaintenance.Items.Item("txtRM").Specific
                                    Call DateBind() : ComboBoxBind() : ConnectToCompany() : oCount() : CreateFolder() ': AddChooseFromList()  ': Matrix_Col() : 

                                    'txt_C.DataBind.SetBound(True, "", "EditDS")
                                    'txtRM.DataBind.SetBound(True, "", "EditDS")
                                    'bt_Search.ChooseFromListUID = "CardCode"
                                    frmMaintenance.AutoManaged = True
                                    'frmMaintenance.SupportedModes = 4
                                    dgCOLUMNS.Columns.Item("Origin").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                                    dgCOLUMNS.Columns.Item("Destination").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                                    dgCOLUMNS.Columns.Item("Service Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox


                                    dgCOLUMNSOTHER.Columns.Item("ItemCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                                    dgCOLUMNSOTHER.Columns.Item("Description").Editable = False

                                    frmMaintenance.EnableMenu("1294", True)
                                    'frmMaintenance.EnableMenu("773", True)
                                    Call Fill_ComboBox()

                                    frmMaintenance.Items.Item("txtRM").Enabled = False
                                    frmMaintenance.Items.Item("txt_Code").Enabled = False

                                    frmMaintenance.Items.Item("txtRM").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                    'SAP_APP.FormItem("tLoading").Enabled = False
                                    frmMaintenance.Items.Item("tLoading").Visible = False
                                    frmMaintenance.Items.Item("txt_T").Enabled = False
                                    frmMaintenance.Items.Item("txt_AppD").Enabled = False
                                    frmMaintenance.Items.Item("txt_AppT").Enabled = False
                                    frmMaintenance.Items.Item("txt_AppN").Enabled = False

                                    fldr_FCR = frmMaintenance.Items.Item("FCR").Specific
                                    fldr_FCR.Select()

                                    frmMaintenance.Freeze(False)
                            End Select
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            txt_AppT = frmMaintenance.Items.Item("txt_AppT").Specific
                            txt_AppN = frmMaintenance.Items.Item("txt_AppN").Specific
                            txt_AppD = frmMaintenance.Items.Item("txt_AppD").Specific
                            txt_T = frmMaintenance.Items.Item("txt_T").Specific

                            Try
                                Select Case pVal.ItemUID
                                    Case "cmd_Delete"
                                        oItems()
                                        If txtRM.Value = "" Then
                                        Else
                                            If SAP_APP.SBO_Application.MessageBox("Are you sure you want to delete this Rate Matrix?", 2, , "Cancel") = 1 Then
                                                Delete_Matrix()
                                            End If
                                        End If
                                    Case "bt_Save"
                                        frmMaintenance.Freeze(True)

                                        txtRM = frmMaintenance.Items.Item("txtRM").Specific
                                        txt_oCount = frmMaintenance.Items.Item("txt_oCount").Specific
                                        txt_Code = frmMaintenance.Items.Item("txt_Code").Specific
                                        cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
                                        cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific

                                        If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                            If txtRM.Value <> "" And txt_Code.Value <> "" Then
                                                If SAP_APP.SBO_Application.MessageBox("If this rate has no record for Other Charges " & vbCrLf & " Please press ""OK"" for New Document" & vbCrLf & "Otherwise " & vbCrLf & "Press ""Update"" to Update Document", 2, "OK", "UPDATE", ) = 2 Then
                                                    If UpdateTo() Then

                                                    End If

                                                Else

                                                    If SAP_APP.SBO_Application.MessageBox("Do you want to save it?", 2, "Yes", "No", ) = 1 Then
                                                        If SaveTo() Then
                                                            frmMaintenance.Items.Item("bt_New").Visible = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf cmb_Format.Value = "" Or cmb_LoadT.Value = "" Then

                                            SAP_APP.SetMessage("Please Choose Format/Load Type!", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                        Else

                                            'If txtRM.Value = "" Then
                                            If txtRM.Value <> "" And txt_Code.Value <> "" Then

                                                If SAP_APP.SBO_Application.MessageBox("Press OK for New Document" & vbCrLf & "Press Update to Update Document", 2, "OK", "UPDATE", ) = 2 Then
                                                    If UpdateTo() = True Then
                                                        frmMaintenance.Items.Item("dt_Start").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                                    End If

                                                Else

                                                    If SAP_APP.SBO_Application.MessageBox("Do you want to save it?", 2, "Yes", "No", ) = 1 Then
                                                        If SaveTo() Then
                                                            frmMaintenance.Items.Item("bt_New").Visible = True
                                                        End If
                                                    End If

                                                End If

                                            Else

                                                If SAP_APP.SBO_Application.MessageBox("Do you want to save it?", 2, "Yes", "No", ) = 1 Then
                                                    If SaveTo() Then
                                                        frmMaintenance.Items.Item("bt_New").Visible = True
                                                    End If
                                                End If

                                            End If
                                        End If
                                        'FOLDERS
                                    Case "3"
                                        'If SAP_APP.SBO_Application.MessageBox("Leaving this will clear your previous Data." & vbCrLf & "Do you want to continue?", 2, "Yes", "No", ) = 1 Then
                                        'Clear()
                                        'frmMaintenance.Freeze(True)
                                        SAP_APP.SAP_Form.PaneLevel = 1
                                        txt_AppT.Value = "1"
                                        txt_AppN.Value = "App_OSEA"
                                        txt_AppD.Value = "App_SEA1"
                                        txt_T.Value = "S"

                                        Try

                                            Try
                                                dtCOLUMNS.Columns.Add("0-5 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                dtCOLUMNS.Columns.Add("6-49 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                dtCOLUMNS.Columns.Add("50-249 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                dtCOLUMNS.Columns.Add("250-999k KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                dtCOLUMNS.Columns.Add("1000UP", SAPbouiCOM.BoFieldsType.ft_Price)
                                            Catch ex As Exception

                                            End Try
                                            dgCOLUMNSOTHER.Columns.Item(2).Visible = False

                                            dgCOLUMNSOTHER.Columns.Item(6).TitleObject.Caption = "North Mindanao"
                                            dgCOLUMNSOTHER.Columns.Item(7).Visible = True
                                            dgCOLUMNSOTHER.Columns.Item(8).Visible = True
                                            dgCOLUMNSOTHER.Columns.Item(9).Visible = True
                                            dgCOLUMNS.Columns.Item(4).Visible = False
                                            dgCOLUMNS.Columns.Item(5).Visible = False
                                            dgCOLUMNS.Columns.Item(6).Visible = False
                                            dgCOLUMNS.Columns.Item(7).Visible = False
                                            dgCOLUMNS.Columns.Item(8).Visible = False

                                            dgCOLUMNS.Columns.Item("5").Visible = True 'Para mag ka error :P
                                        Catch Exp As Exception

                                            Try

                                                oItems()

                                                For i As Integer = cmb_Format.ValidValues.Count - 1 To 0 Step -1
                                                    cmb_Format.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                Next
                                                For i As Integer = cmb_LoadT.ValidValues.Count - 1 To 0 Step -1
                                                    cmb_LoadT.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                Next

                                                If txt_AppT.Value = "1" Then
                                                    cmb_Format.ValidValues.Add("1", "Shipping Line Tariff Rate")

                                                    cmb_LoadT.ValidValues.Add("LCL", "Less Container Load")
                                                    cmb_LoadT.ValidValues.Add("FCL", "Full Container Load")

                                                    If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                                        cmb_Format.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                        cmb_LoadT.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                    End If

                                                    frmMaintenance.Items.Item("cmb_Format").Refresh()
                                                    frmMaintenance.Items.Item("cmb_Format").DisplayDesc = False
                                                    frmMaintenance.Items.Item("cmb_LoadT").Refresh()
                                                    frmMaintenance.Items.Item("cmb_LoadT").DisplayDesc = False
                                                End If

                                                frmMaintenance.Items.Item("cmb_LoadT").DisplayDesc = True
                                                frmMaintenance.Items.Item("cmb_Format").DisplayDesc = True

                                                If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                                    frmMaintenance.Items.Item("dg_Less").Visible = True
                                                    frmMaintenance.Items.Item("dg_Others").Visible = False
                                                End If

                                            Catch ex As Exception
                                                SAP_APP.SetMessage(ex)
                                            End Try
                                        End Try
                                        'GoTo Unfreeze

                                        'End If
                                    Case "4"

                                        oItems()

                                        If txtRM.Value = "" Or txt_Code.Value = "" Then
                                            Clear()
                                            GoTo oOo
                                        Else
                                            If SAP_APP.SBO_Application.MessageBox("Leaving this Form will clear your previous Data." & vbCrLf & "Do you want to continue?", 2, "Yes", "No", ) = 1 Then
                                                'Clear()
                                                oDgClear()
oOo:
                                                frmMaintenance.Freeze(True)
                                                SAP_APP.SAP_Form.PaneLevel = 2
                                                txt_AppT.Value = "2"
                                                txt_T.Value = "L"

                                                fldr_FCR = frmMaintenance.Items.Item("FCR").Specific
                                                fldr_FCR.Select()

                                                Try
                                                    dgCOLUMNSOTHER.Columns.Item(2).Visible = True

                                                    dgCOLUMNSOTHER.Columns.Item(6).TitleObject.Caption = "Mindanao"
                                                    dgCOLUMNSOTHER.Columns.Item(7).Visible = False
                                                    dgCOLUMNSOTHER.Columns.Item(8).Visible = False
                                                    dgCOLUMNSOTHER.Columns.Item(9).Visible = False
                                                    dgCOLUMNS.Columns.Item(4).Visible = True
                                                    dgCOLUMNS.Columns.Item(5).Visible = True
                                                    dgCOLUMNS.Columns.Item(6).Visible = False
                                                    dgCOLUMNS.Columns.Item(7).Visible = False
                                                    dgCOLUMNS.Columns.Item(8).Visible = False

                                                    dgCOLUMNS.Columns.Item("5").Visible = True 'Para mag ka error :P
                                                Catch Exp As Exception
                                                    Try

                                                        oItems()

                                                        For i As Integer = cmb_Format.ValidValues.Count - 1 To 0 Step -1
                                                            cmb_Format.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                        Next
                                                        For i As Integer = cmb_LoadT.ValidValues.Count - 1 To 0 Step -1
                                                            cmb_LoadT.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                        Next

                                                        If txt_AppT.Value = "2" Then
                                                            cmb_Format.ValidValues.Add("1", "Less Truck Load")
                                                            cmb_Format.ValidValues.Add("2", "Full Truck Load")

                                                            cmb_LoadT.ValidValues.Add("PKR-3", "Per Kilo Rate - 3KGS")
                                                            cmb_LoadT.ValidValues.Add("PKR-5", "Per Kilo Rate - 5KGS")
                                                            cmb_LoadT.ValidValues.Add("FRT", "Fixed Rate")
                                                            cmb_LoadT.ValidValues.Add("FTL", "Full Truck Load")

                                                            If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                                                cmb_Format.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                                cmb_LoadT.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                            End If

                                                            frmMaintenance.Items.Item("cmb_Format").Refresh()
                                                            frmMaintenance.Items.Item("cmb_Format").DisplayDesc = False
                                                            frmMaintenance.Items.Item("cmb_LoadT").Refresh()
                                                            frmMaintenance.Items.Item("cmb_LoadT").DisplayDesc = False

                                                        End If
                                                        frmMaintenance.Items.Item("cmb_LoadT").DisplayDesc = True
                                                        frmMaintenance.Items.Item("cmb_Format").DisplayDesc = True

                                                        If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                                            frmMaintenance.Items.Item("dg_Less").Visible = True
                                                            frmMaintenance.Items.Item("dg_Others").Visible = False
                                                        End If

                                                    Catch ex As Exception
                                                        SAP_APP.SetMessage(ex)
                                                    End Try
                                                End Try
                                                GoTo Unfreeze
                                                'Else
                                                'frmMaintenance.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                            End If
                                        End If
                                    Case "5"
                                        If txtRM.Value = "" Or txt_Code.Value = "" Then
                                            Clear()
                                            GoTo oO
                                        Else
                                            If SAP_APP.SBO_Application.MessageBox("Leaving this will clear your previous Data." & vbCrLf & "Do you want to continue?", 2, "Yes", "No", ) = 1 Then
                                                'Clear()
                                                oDgClear()
oO:
                                                frmMaintenance.Freeze(True)
                                                SAP_APP.SAP_Form.PaneLevel = 3
                                                txt_AppT.Value = "3"
                                                txt_T.Value = "A"
                                                Try
                                                    dtCOLUMNS.Columns.Add("0-5 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                    'dtCOLUMNS.Columns.Add("6-49 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                    'dtCOLUMNS.Columns.Add("50-249 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                    dtCOLUMNS.Columns.Add("250-499 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                    dtCOLUMNS.Columns.Add("500-999 KG", SAPbouiCOM.BoFieldsType.ft_Price)
                                                    'dtCOLUMNS.Columns.Add("1000UP", SAPbouiCOM.BoFieldsType.ft_Price)
                                                Catch Exp As Exception
                                                    dgCOLUMNSOTHER.Columns.Item(2).Visible = True

                                                    dgCOLUMNSOTHER.Columns.Item(6).TitleObject.Caption = "Mindanao"
                                                    dgCOLUMNSOTHER.Columns.Item(7).Visible = False
                                                    dgCOLUMNSOTHER.Columns.Item(8).Visible = False
                                                    dgCOLUMNSOTHER.Columns.Item(9).Visible = False
                                                    dgCOLUMNS.Columns.Item(4).Visible = True
                                                    dgCOLUMNS.Columns.Item(5).Visible = True
                                                    dgCOLUMNS.Columns.Item(6).Visible = True
                                                    dgCOLUMNS.Columns.Item(7).Visible = True
                                                    dgCOLUMNS.Columns.Item(8).Visible = True
                                                    GoTo NoError
                                                End Try
NoError:
                                                Try
                                                    oItems()

                                                    For i As Integer = cmb_Format.ValidValues.Count - 1 To 0 Step -1
                                                        cmb_Format.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                    Next
                                                    For i As Integer = cmb_LoadT.ValidValues.Count - 1 To 0 Step -1
                                                        cmb_LoadT.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                                    Next

                                                    If txt_AppT.Value = "3" Then
                                                        cmb_Format.ValidValues.Add("1", "Airline Tariff Rate")

                                                        cmb_LoadT.ValidValues.Add("ATR", "Cebu Pacific")
                                                        cmb_LoadT.ValidValues.Add("ATR2", "PAL")
                                                        cmb_LoadT.ValidValues.Add("PKR-3", "Per Kilo Rate - 3KGS")
                                                        cmb_LoadT.ValidValues.Add("PKR-5", "Per Kilo Rate - 5KGS")
                                                        cmb_LoadT.ValidValues.Add("PKR-10", "Per Kilo Rate - 10KGS")
                                                        cmb_LoadT.ValidValues.Add("SME", "SME Rate")
                                                        cmb_LoadT.ValidValues.Add("FRT", "Fixed Rate")
                                                        cmb_LoadT.ValidValues.Add("PBX", "Per Box")

                                                        If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                                            cmb_Format.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                            cmb_LoadT.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                        End If
                                                    End If

                                                    frmMaintenance.Items.Item("cmb_Format").Refresh()
                                                    frmMaintenance.Items.Item("cmb_Format").DisplayDesc = True
                                                    frmMaintenance.Items.Item("cmb_LoadT").Refresh()
                                                    frmMaintenance.Items.Item("cmb_LoadT").DisplayDesc = True
                                                Catch ex As Exception
                                                    SAP_APP.SetMessage(ex)
                                                End Try
                                                GoTo Unfreeze
                                            End If
                                        End If
                                    Case "FCR"

                                        If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                            frmMaintenance.Items.Item("dg_Others").Visible = False
                                            frmMaintenance.Items.Item("dg_Less").Visible = True
                                        End If

                                        cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
                                        cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific

                                        frmMaintenance.Items.Item("dg_Others").Visible = False

                                        If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                            frmMaintenance.Items.Item("dg_Less").Visible = True
                                        End If

                                    Case "OC"

                                        frmMaintenance.Items.Item("dg_Less").Visible = False

                                        cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
                                        cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific

                                        frmMaintenance.Items.Item("dg_Less").Visible = False

                                        If cmb_Format.Value <> "" And cmb_LoadT.Value <> "" Then
                                            frmMaintenance.Items.Item("dg_Others").Visible = True
                                        End If
                                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    Case "bt_Line"
                                        Try
                                            If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                                oGrid = frmMaintenance.Items.Item("dg_Others").Specific
                                                oGrid.DataTable.Rows.Add(1)
                                            ElseIf frmMaintenance.Items.Item("dg_Less").Visible = True Then
                                                oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                                                oGrid.DataTable.Rows.Add(1)
                                            End If
                                        Catch ex As Exception
                                            'SAP_APP.SetMessage("This is an ERROR. Please choose LoadType First", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End Try
                                    Case "bt_Count"
                                        Try
                                            oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                                            MsgBox(oGrid.DataTable.Rows.Count)
                                        Catch ex As Exception
                                            SAP_APP.SetMessage(ex)
                                        End Try
                                    Case "bt_Delete"
                                        Try
                                            frmMaintenance.Freeze(True)
                                            If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                                oGrid = frmMaintenance.Items.Item("dg_Others").Specific
                                                If oGrid.Rows.Count > 0 Then
                                                    For i As Integer = 1 To oGrid.Rows.Count - 1
                                                        If oGrid.Rows.IsSelected(i) = True Then
                                                            oGrid.DataTable.Rows.Remove(i)
                                                            GoTo x
                                                        End If
                                                    Next
                                                End If

                                            ElseIf frmMaintenance.Items.Item("dg_Less").Visible = True Then
                                                oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                                                If oGrid.Rows.Count > 0 Then
                                                    For i As Integer = 1 To oGrid.Rows.Count - 1
                                                        If oGrid.Rows.IsSelected(i) = True Then
                                                            oGrid.DataTable.Rows.Remove(i)
                                                            GoTo x
                                                        End If
                                                    Next
                                                End If

                                            End If
x:
                                            frmMaintenance.Freeze(False)
                                        Catch ex As Exception
                                            oGrid.DataTable.Rows.Clear()
                                        End Try
                                    Case "bt_New"
                                        If Clear() = True Then
                                            frmMaintenance.Items.Item("bt_New").Visible = False
                                        End If
                                    Case "bt_Search"
                                        'SAP_APP.SetMessage("UNDER MAINTENANCE..")
                                    Case "1"

                                End Select

                            Catch ex As Exception
                                SAP_APP.SetMessage("Please SAVE your document before proceeding to other tab.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                If frmMaintenance.Items.Item("dg_Less").Visible = True Then
                                    frmMaintenance.Items.Item("FCR").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                ElseIf frmMaintenance.Items.Item("dg_Others").Visible = True Then
                                    frmMaintenance.Items.Item("OC").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                End If

                            End Try

                    End Select
            End Select
            'This Line is for Formatted Search----------------------------------------------------------------------'
            'Select Case pVal.BeforeAction                                                                           '
            '    Case False                                                                                          '
            '    Case True                                                                                           '

            '        Select Case pVal.ItemUID                                                                        '
            '            Case "1"                                                                                    '
            '                Select Case pVal.EventType                                                              '
            '                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED                                        '
            '                        Select Case frmMaintenance.Mode                                                 '
            '                            Case SAPbouiCOM.BoFormMode.fm_FIND_MODE                                     '
            '                                If CheckRM() = True Then                                                '
            '                                    frmMaintenance.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE          '
            '                                    frmMaintenance.Items.Item("txtRM").Enabled = False                  '
            '                                    frmMaintenance.Items.Item("dt_Start").Click(SAPbouiCOM.BoCellClickType.ct_Double)
            '                                    frmMaintenance.Items.Item("txt_Code").Enabled = False               '
            '                                End If                                                                  '
            '                                BubbleEvent = False                                                     '
            '                        End Select                                                                      '
            '                End Select                                                                              '
            '        End Select                                                                                      '
            'End Select                                                                                              '
            '-------------------------------------------------------------------------------------------------------'
        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
Unfreeze:
        frmMaintenance.Freeze(False)
    End Sub

#End Region

#Region "Date and Binding"
    Private Sub DateBind()
        Try

            frmMaintenance.DataSources.UserDataSources.Add("EditDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            frmMaintenance.DataSources.UserDataSources.Add("txt_VMW", SAPbouiCOM.BoDataType.dt_PRICE)
            frmMaintenance.DataSources.UserDataSources.Add("txt_Charge", SAPbouiCOM.BoDataType.dt_PRICE)
            frmMaintenance.DataSources.UserDataSources.Add("dt_Start", SAPbouiCOM.BoDataType.dt_DATE)
            frmMaintenance.DataSources.UserDataSources.Add("dt_End", SAPbouiCOM.BoDataType.dt_DATE)

            Dim dt_Start As SAPbouiCOM.EditText = frmMaintenance.Items.Item("dt_Start").Specific
            Dim dt_End As SAPbouiCOM.EditText = frmMaintenance.Items.Item("dt_End").Specific
            txt_VMW = frmMaintenance.Items.Item("txt_VMW").Specific
            txt_Charge = frmMaintenance.Items.Item("txt_Charge").Specific

            'txt_VMW.DataBind.SetBound(True, "", "txt_VMW")
            'txt_Charge.DataBind.SetBound(True, "", "txt_Charge")
            dt_Start.DataBind.SetBound(True, "", "dt_Start")
            dt_End.DataBind.SetBound(True, "", "dt_End")

            dt_Start.Value = Format(Date.Today, "yyyyMMdd")
            dt_End.Value = Format(Date.Today, "yyyyMMdd")

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Sub
#End Region

#Region "Combobox"
    Private Sub ComboBoxBind()

    End Sub
#End Region

#Region "ConnectToCompany"
    Private Sub ConnectToCompany()
        Try

            oCompany = SBO_Application.Company.GetDICompany

        Catch ex As Exception
            frmMaintenance.Freeze(False)
            'SAP_APP.SetMessage(ex)
        End Try
    End Sub
#End Region

#Region "oCount"
    Private Sub oCount()
        Try
            Dim oCount As SAPbouiCOM.EditText = frmMaintenance.Items.Item("txt_oCount").Specific
            'conString = "Data Source=" & oCompany.Server & ";Initial Catalog=" & oCompany.CompanyDB & ";Integrated Security=True"
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            txt_AppN = frmMaintenance.Items.Item("txt_AppN").Specific

            oQuery = "SELECT TOP 1 AppNo FROM [" & oCompany.CompanyDB & "]..[" & txt_AppN.Value & "] ORDER BY AppId DESC "
            oRec.DoQuery(oQuery)
            'oConnect = New SqlConnection(conString)
            'oConnect.Open()
            'oCommand = New SqlCommand(oQuery, oConnect)

            Dim i As String = Mid(oRec.Fields.Item(0).Value, 4)
            Dim ii As Integer = CInt(i) + 1
            oCount.Value = String.Format("{0:D5}", ii)
            'oConnect.Close()
            'MsgBox(oCount.Value)

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Sub
#End Region

#Region "Parameters"
    Private Sub Parameters()

        oItems()

        Try
            txt_Code.Value = txtRM.Value & "-" & txt_T.Value & "-" & cmb_Format.Selected.Value & "-" & cmb_LoadT.Selected.Value
            'With oCommand.Parameters
            '    .AddWithValue("@txtRM", txtRM.Value)
            '    .AddWithValue("@txt_AppT", txt_AppT.Value)
            '    .AddWithValue("@dt_Start", dt_Start.Value)
            '    .AddWithValue("@dt_End", dt_End.Value)
            '    .AddWithValue("@txt_Code", txt_Code.Value)
            '    .AddWithValue("@cmb_Format", cmb_Format.Selected.Description)
            '    .AddWithValue("@cmb_LoadT", cmb_LoadT.Selected.Description)
            '    .AddWithValue("@txt_VMW", txt_VMW.Value)
            '    .AddWithValue("@txt_Charge", txt_Charge.Value)
            '    '.AddWithValue("@txt_NumB", txt_NumB.Value)
            '    '.AddWithValue("@txt_RateB", txt_RateB.Value)
            'End With
        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Sub
#End Region

#Region "Choose From List"
    Private Sub AddChooseFromList()
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = frmMaintenance.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CardCode"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)


        Catch ex As Exception
            SAP_APP.SetMessage(ex, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "CreateFolder"
    Private Sub CreateFolder()
        Try

            Dim oItem As SAPbouiCOM.Item
            Dim oFolder As SAPbouiCOM.Folder

            oItem = frmMaintenance.Items.Add("FCR", SAPbouiCOM.BoFormItemTypes.it_FOLDER)

            oItem.Left = 7
            oItem.Width = 320
            oItem.Top = 167
            oItem.Height = 20

            oItem.FromPane = 1
            oItem.ToPane = 3

            oFolder = oItem.Specific
            'oFolder.Pane = 1
            oFolder.Caption = "Freight Charges Rates"
            oFolder.DataBind.SetBound(True, "", "FolderDS")
            oFolder.Select()

            oItem = frmMaintenance.Items.Add("OC", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Left = 165
            oItem.Width = 320
            oItem.Top = 167
            oItem.Height = 20

            oItem.FromPane = 1
            oItem.ToPane = 3

            oFolder = oItem.Specific
            'oFolder.Pane = 1
            oFolder.Caption = "Other Charges"
            oFolder.DataBind.SetBound(True, "", "FolderDS")
            oFolder.GroupWith("FCR")

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Sub
#End Region

#Region "Items"
    Private Function oItems() As Boolean


        txt_oCount = frmMaintenance.Items.Item("txt_oCount").Specific
        txt_T = frmMaintenance.Items.Item("txt_T").Specific
        txtRM = frmMaintenance.Items.Item("txtRM").Specific
        txt_Code = frmMaintenance.Items.Item("txt_Code").Specific
        cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
        cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific
        oGrid = frmMaintenance.Items.Item("dg_Less").Specific
        dt_Start = frmMaintenance.Items.Item("dt_Start").Specific
        dt_End = frmMaintenance.Items.Item("dt_End").Specific
        'txt_NumB = frmMaintenance.Items.Item("txt_NumB").Specific
        'txt_RateB = frmMaintenance.Items.Item("txt_RateB").Specific
        dg_Others = frmMaintenance.Items.Item("dg_Others").Specific
        txt_AppT = frmMaintenance.Items.Item("txt_AppT").Specific
        txt_AppN = frmMaintenance.Items.Item("txt_AppN").Specific
        txt_VMW = frmMaintenance.Items.Item("txt_VMW").Specific
        txt_Charge = frmMaintenance.Items.Item("txt_Charge").Specific
        txt_Charg2 = frmMaintenance.Items.Item("txt_Charg2").Specific
        txt_Peri = frmMaintenance.Items.Item("txt_Peri").Specific
        txt_GMA = frmMaintenance.Items.Item("txt_GMA").Specific
        oBP = frmMaintenance.Items.Item("41").Specific

    End Function
#End Region

    Private Function Delete_Matrix() As Boolean
        Try

            oItems()

            If oCompany.InTransaction = False Then
                oCompany.StartTransaction()
            End If

            Try
                oQuery = "DELETE FROM [" & oCompany.CompanyDB & "]..APP_OSEA WHERE AppNo = '" & txtRM.Value & "' " & _
                    "DELETE FROM [" & oCompany.CompanyDB & "]..APP_SEA1 WHERE AppNo = '" & txtRM.Value & "'"
                oRec.DoQuery(oQuery)
            Catch ex As Exception
                SAP_APP.SetMessage(ex)
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Return False
            End Try

            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

            dgCOLUMNS.DataTable.ExecuteQuery("SELECT TOP 0 FROM [" & oCompany.CompanyDB & "]..APP_SEA1")
            txtRM.Value = ""

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return False
        End Try
    End Function

    Private Function SaveTo() As Boolean
        SaveTo = True
        Dim oCounts As SAPbouiCOM.EditText = frmMaintenance.Items.Item("txt_oCount").Specific

        Try

            oItems()

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If txtRM.Value = "" Then

                txtRM.Value = "RM-1" & txt_oCount.Value & ""

                Call Parameters()
                oQuery = ("EXEC [APP_DB]..sp_InsertHeader " & oCompany.CompanyDB & ", " & txt_AppN.Value & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', " & _
                            "'" & dt_Start.Value & "', '" & dt_End.Value & "', '" & IIf(txt_VMW.Value = "", "0", txt_VMW.Value) & "', '" & IIf(txt_Charge.Value = "", "0", txt_Charge.Value) & "', '" & IIf(txt_Charg2.Value = "", "0", txt_Charg2.Value) & "', '" & IIf(txt_Peri.Value = "", "0", txt_Peri.Value) & "', '" & IIf(txt_GMA.Value = "", "0", txt_GMA.Value) & "', '" & Replace(oBP.Value, "'", "") & "', '1'")
                oRec.DoQuery(oQuery)
                GoTo Save

            Else

                Call Parameters()
                oQuery = ("EXEC [APP_DB]..sp_InsertHeader " & oCompany.CompanyDB & ", " & txt_AppN.Value & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', " & _
                            "'" & dt_Start.Value & "', '" & dt_End.Value & "', '" & IIf(txt_VMW.Value = "", "0", txt_VMW.Value) & "', '" & IIf(txt_Charge.Value = "", "0", txt_Charge.Value) & "', '" & IIf(txt_Charg2.Value = "", "0", txt_Charg2.Value) & "', '" & IIf(txt_Peri.Value = "", "0", txt_Peri.Value) & "', '" & IIf(txt_GMA.Value = "", "0", txt_GMA.Value) & "', '" & Replace(oBP.Value, "'", "") & "', '1'")
                oRec.DoQuery(oQuery)

                If txtRM.Value <> "" And txt_Code.Value <> "" Then


Save:
                    If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = dg_Others.DataTable

                        For oRow As Integer = 0 To oDT.Rows.Count - 1
                            Select Case txt_AppT.Value
                                Case "1"
                                    Call Parameters()
                                    oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '" & oDT.GetValue("ItemCode", oRow) & "', " & _
                                                " '" & oDT.GetValue("Description", oRow) & "', '" & oDT.GetValue("1x10", oRow) & "', '" & oDT.GetValue("1x20", oRow) & "', '" & oDT.GetValue("AccountCode", oRow) & "', '" & oDT.GetValue("Luzon", oRow) & "', '" & oDT.GetValue("Visayas", oRow) & "', '" & oDT.GetValue("Mindanao", oRow) & "', '" & oDT.GetValue("South Mindanao", oRow) & "' , '2'"
                                    oRec.DoQuery(oQuery)
                                Case "2"
                                    Call Parameters()
                                    oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '" & oDT.GetValue("ItemCode", oRow) & "', " & _
                                                " '" & oDT.GetValue("Description", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '0', '" & oDT.GetValue("AccountCode", oRow) & "', '" & oDT.GetValue("Luzon", oRow) & "', '" & oDT.GetValue("Visayas", oRow) & "', '" & oDT.GetValue("Mindanao", oRow) & "', '-1' , '3'"
                                    oRec.DoQuery(oQuery)
                            End Select
                        Next

                    Else

                        oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = oGrid.DataTable

                        If txt_AppT.Value <> "3" Then

                            For oRow As Integer = 0 To oDT.Rows.Count - 1
                                Call Parameters()
                                Select Case txt_AppT.Value
                                    Case "2"
                                        oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '" & oDT.GetValue(4, oRow) & "', '" & oDT.GetValue(5, oRow) & "', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '5'"
                                        oRec.DoQuery(oQuery)
                                    Case "1"
                                        oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '-1', '-1', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '-1', '4'"
                                        oRec.DoQuery(oQuery)
                                End Select

                            Next

                        Else

                            For oRow As Integer = 0 To oDT.Rows.Count - 1
                                Call Parameters()
                                oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '" & oDT.GetValue(4, oRow) & "', '" & oDT.GetValue(5, oRow) & "', '" & oDT.GetValue(6, oRow) & "', '" & oDT.GetValue(7, oRow) & "', '" & oDT.GetValue(8, oRow) & "', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '-1', '6'"
                                oRec.DoQuery(oQuery)
                            Next

                        End If

                    End If
                End If
            End If

            Call oCount()
            UpdateTo()
        Catch ex As Exception
            frmMaintenance.Freeze(False)
            SAP_APP.SetMessage(ex)
        End Try
    End Function

    Private Function UpdateTo() As Boolean
        Try

            oItems()

            Dim oQuery As String
            Dim oCounts As SAPbouiCOM.EditText = frmMaintenance.Items.Item("txt_oCount").Specific

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oCompany.StartTransaction()

            oQuery = "UPDATE [" & oCompany.CompanyDB & "]..[" & txt_AppN.Value & "]" & _
                    "SET StartDate = '" & dt_Start.Value & "', ExpiryDate = '" & dt_End.Value & "', VMW = " & IIf(txt_VMW.Value = "", 0, txt_VMW.Value) & ", Freight = " & IIf(txt_Charge.Value = "", 0, txt_Charge.Value) & ", Freigh2 = " & IIf(txt_Charg2.Value = "", 0, txt_Charg2.Value) & ", isGMA = " & IIf(txt_GMA.Value = "", 0, txt_GMA.Value) & ", [BP Name] = '" & Replace(oBP.Value, "'", "") & "' " & _
                    "WHERE AppNo = '" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "' "
            oRec.DoQuery(oQuery)
            If frmMaintenance.Items.Item("dg_Others").Visible = True Then
                oQuery = "DELETE FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                                "WHERE AppNo = '" & txtRM.Value & "' AND Code = '" & txt_Code.Value & "' AND LoadType = '-1'"
                oRec.DoQuery(oQuery)

                Dim oDT As SAPbouiCOM.DataTable
                oDT = dg_Others.DataTable
                For oRow As Integer = 0 To oDT.Rows.Count - 1
                    Select Case txt_AppT.Value
                        Case "1"
                            Call Parameters()

                            Dim oMindanao As Decimal
                            Try
                                oMindanao = oDT.GetValue("Mindanao", oRow)
                            Catch ex As Exception
                                oMindanao = oDT.GetValue("North Mindanao", oRow)
                            End Try
                            oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '" & oDT.GetValue("ItemCode", oRow) & "', " & _
                                                " '" & oDT.GetValue("Description", oRow) & "', '" & oDT.GetValue("1x10", oRow) & "', '" & oDT.GetValue("1x20", oRow) & "', '" & oDT.GetValue("AccountCode", oRow) & "', '" & oDT.GetValue("Luzon", oRow) & "', '" & oDT.GetValue("Visayas", oRow) & "', '" & oMindanao & "', '" & oDT.GetValue("South Mindanao", oRow) & "' , '2'"
                            oRec.DoQuery(oQuery)
                        Case Else
                            Call Parameters()
                            oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '" & oDT.GetValue("ItemCode", oRow) & "', " & _
                                                " '" & oDT.GetValue("Description", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "' , '0', '" & oDT.GetValue("AccountCode", oRow) & "', '" & oDT.GetValue("Luzon", oRow) & "', '" & oDT.GetValue("Visayas", oRow) & "', '" & oDT.GetValue("Mindanao", oRow) & "', '-1' , '3'"
                            oRec.DoQuery(oQuery)
                    End Select
                Next

            Else

                If txt_AppT.Value <> "3" Then

                    oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                    oQuery = "DELETE FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                "WHERE AppNo = '" & txtRM.Value & "' AND Code = '" & txt_Code.Value & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "'"
                    oRec.DoQuery(oQuery)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = oGrid.DataTable
                    For oRow As Integer = 0 To oDT.Rows.Count - 1

                        Select Case txt_AppT.Value
                            Case "2"
                                Call Parameters()
                                oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '" & oDT.GetValue(4, oRow) & "', '" & oDT.GetValue(5, oRow) & "', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '5'"
                                oRec.DoQuery(oQuery)
                            Case "1"
                                Call Parameters()
                                oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '-1', '-1', '-1', '-1', '-1', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '-1', '4'"
                                oRec.DoQuery(oQuery)
                        End Select

                    Next

                Else

                    oGrid = frmMaintenance.Items.Item("dg_Less").Specific
                    oRec.DoQuery("DELETE FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                "WHERE AppNo = '" & txtRM.Value & "' AND Code = '" & txt_Code.Value & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "'")
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = oGrid.DataTable
                    For oRow As Integer = 0 To oDT.Rows.Count - 1

                        Call Parameters()
                        oQuery = "EXEC [APP_DB]..sp_InsertDetails " & oCompany.CompanyDB & ", '" & oCounts.Value & "', '" & txtRM.Value & "', '" & txt_AppT.Value & "', '" & txt_Code.Value & "', '" & cmb_Format.Selected.Description & "', '" & cmb_LoadT.Selected.Description & "', '" & oDT.GetValue("Origin", oRow) & "', " & _
                                            " '" & oDT.GetValue("Destination", oRow) & "', '" & oDT.GetValue("Service Type", oRow) & "', '" & oDT.GetValue("Rate", oRow) & "', '" & oDT.GetValue(4, oRow) & "', '" & oDT.GetValue(5, oRow) & "', '" & oDT.GetValue(6, oRow) & "', '" & oDT.GetValue(7, oRow) & "', '" & oDT.GetValue(8, oRow) & "', '" & txt_AppD.Value & "', '" & oRow + 1 & "', '-1', '-1', '0', '-1', '-1', '-1', '-1', '-1', '-1', '6'"
                        Console.WriteLine(oQuery)
                        oRec.DoQuery(oQuery)
                    Next

                End If

            End If

            UpdateTo = True
            SAP_APP.SetMessage("Successfully Saved!", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            frmMaintenance.Items.Item("dt_Start").Click(SAPbouiCOM.BoCellClickType.ct_Double)
        Catch ex As Exception
            frmMaintenance.Freeze(False)
            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            SAP_APP.SetMessage(ex)
        End Try
    End Function

    Private Sub Fill_ComboBox()
        Try
            'oPrgress = SBO_Application.StatusBar.CreateProgressBar("Retriving Data. Please wait...", 27, True)
            SBO_Application.SetStatusBarMessage("Retriving Data. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            oCombo = dgCOLUMNS.Columns.Item("Origin")
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oRec.DoQuery("SELECT * FROM [" & oCompany.CompanyDB & "]..[@LOCATION]")

            If oRec.RecordCount > 0 Then
                While oRec.EoF = False
                    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(1).Value.ToString)
                    oRec.MoveNext()
                End While
                oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            End If

            oCombo = dgCOLUMNS.Columns.Item("Destination")

            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oRec.DoQuery("SELECT * FROM [" & oCompany.CompanyDB & "]..[@LOCATION]")

            If oRec.RecordCount > 0 Then
                While oRec.EoF = False
                    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(1).Value.ToString)
                    oRec.MoveNext()
                End While
                oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            End If

            cmb_CopyTo = frmMaintenance.Items.Item("cmb_CopyTo").Specific

            For i As Integer = cmb_CopyTo.ValidValues.Count - 1 To 0 Step -1
                cmb_CopyTo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oRec.DoQuery("SELECT AppNo FROM [" & oCompany.CompanyDB & "]..[App_OSEA] GROUP BY AppNo ORDER BY AppNo")

            If oRec.RecordCount > 0 Then
                While oRec.EoF = False
                    cmb_CopyTo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(0).Value.ToString)
                    oRec.MoveNext()
                End While
            End If

            oCombo = dgCOLUMNSOTHER.Columns.Item("ItemCode")

            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oRec.DoQuery("SELECT ItemCode, ItemName, FormatCode FROM [" & oCompany.CompanyDB & "]..[vintel_oitm]")

            If oRec.RecordCount > 0 Then
                While oRec.EoF = False
                    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(1).Value.ToString)
                    oRec.MoveNext()
                    'oPrgress.Value += 1
                End While
            End If

            'oPrgress.Stop()

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Sub

    ReadOnly Property dtCOLUMNS() As SAPbouiCOM.DataTable
        Get
            Return dgCOLUMNS.DataTable
        End Get
    End Property

    ReadOnly Property dgCOLUMNS() As SAPbouiCOM.Grid
        Get
            Return frmMaintenance.Items.Item("dg_Less").Specific
        End Get
    End Property

    ReadOnly Property dtCOLUMNSFULL() As SAPbouiCOM.DataTable
        Get
            Return dgCOLUMNSFULL.DataTable
        End Get
    End Property

    ReadOnly Property dgCOLUMNSFULL() As SAPbouiCOM.Grid
        Get
            Return frmMaintenance.Items.Item("dg_Full").Specific
        End Get
    End Property

    ReadOnly Property dtCOLUMNSOTHER() As SAPbouiCOM.DataTable
        Get
            Return dgCOLUMNSFULL.DataTable
        End Get
    End Property

    ReadOnly Property dgCOLUMNSOTHER() As SAPbouiCOM.Grid
        Get
            Return frmMaintenance.Items.Item("dg_Others").Specific
        End Get
    End Property

    Private Function Clear() As Boolean
        Clear = True
        Try
            frmMaintenance.Freeze(True)

            oItems()

            txtRM.Value = ""
            txt_Code.Value = ""
            dt_Start.Value = Format(Date.Today, "yyyyMMdd")
            dt_End.Value = Format(Date.Today, "yyyyMMdd")
            'txt_NumB.Value = ""
            'txt_RateB.Value = ""

            oGrid = frmMaintenance.Items.Item("dg_Others").Specific
            oGrid.DataTable.Rows.Clear()
            oGrid = frmMaintenance.Items.Item("dg_Full").Specific
            oGrid.DataTable.Rows.Clear()
            oGrid = frmMaintenance.Items.Item("dg_Less").Specific
            oGrid.DataTable.Rows.Clear()

            frmMaintenance.Freeze(False)
        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Function

    Private Function oDgClear() As Boolean
        Try
            oGrid = frmMaintenance.Items.Item("dg_Others").Specific
            oGrid.DataTable.Rows.Clear()
            oGrid = frmMaintenance.Items.Item("dg_Full").Specific
            oGrid.DataTable.Rows.Clear()
            oGrid = frmMaintenance.Items.Item("dg_Less").Specific
            oGrid.DataTable.Rows.Clear()
        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Function

    Private Function CheckRM() As Boolean
        Try
            frmMaintenance.Freeze(True)

            oItems()
            Dim oStr As String
            If txtRM.Value <> "" Then

                CheckRM = True

                oQuery = "SELECT TOP 1 CONVERT(varchar(25),StartDate,112) as StartDate, CONVERT(varchar(25),ExpiryDate,112) as ExpiryDate, VMW, Freight, Freigh2, isGMA, [BP Name] " & _
                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppN.Value & "] WHERE AppNo ='" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "'"
                oRec.DoQuery(oQuery)

                oBP.Value = oRec.Fields.Item(5).Value
                txt_VMW.Value = oRec.Fields.Item("VMW").Value.ToString
                txt_Charge.Value = oRec.Fields.Item("Freight").Value.ToString
                txt_Charg2.Value = oRec.Fields.Item("Freigh2").Value.ToString
                txt_GMA.Value = oRec.Fields.Item("isGMA").Value.ToString
                dt_Start.Value = oRec.Fields.Item(0).Value.ToString
                dt_End.Value = oRec.Fields.Item(1).Value.ToString

                If frmMaintenance.Items.Item("dg_Others").Visible = True Then

                    Select Case txt_AppT.Value
                        Case "1"
                            oStr = "SELECT ItemCode, Description, AccountCode, Luzon, Visayas, Mindanao [North Mindanao], [South Mindanao], Price [1x10], Rate2 [1x20] " & _
                                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                            "WHERE Code = '" & txt_Code.Value & "' AND AppNo = '" & txtRM.Value & "' AND LoadType = '-1' AND AppType = '" & txt_AppT.Value & "' AND CompanyId = '" & oCompany.CompanyDB & "'"
                            dg_Others.DataTable.ExecuteQuery(oStr)

                        Case Else
                            oStr = "SELECT ItemCode, Description, Price [Rate], AccountCode, Luzon, Visayas, Mindanao " & _
                                                                        "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                                                        "WHERE Code = '" & txt_Code.Value & "' AND AppNo = '" & txtRM.Value & "' AND LoadType = '-1' AND AppType = '" & txt_AppT.Value & "' AND CompanyId = '" & oCompany.CompanyDB & "'"
                            dg_Others.DataTable.ExecuteQuery(oStr)

                    End Select

                    dg_Others.Columns.Item("ItemCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    dg_Others.Columns.Item("Description").Editable = False
                    Dim oText As SAPbouiCOM.EditTextColumn = dg_Others.Columns.Item("AccountCode")
                    oText.LinkedObjectType = "1"

                    Fill_ComboBox()

                Else
                    If txt_T.Value <> "A" Then

                        oStr = "SELECT Origin, Destination, SerType1 as [Service Type], SerType2 as Rate " & IIf(txt_T.Value = "L", ", o1st [0-5 KG], o2nd [6-49 KG]", "") & "" & _
                                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                            "WHERE AppNo ='" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "' AND CompanyId = '" & oCompany.CompanyDB & "'"


                        dgCOLUMNS.DataTable.ExecuteQuery(oStr)
                        oRec.DoQuery("SELECT Code " & _
                                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                            "WHERE AppNo ='" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "' AND CompanyId = '" & oCompany.CompanyDB & "'")
                        txt_Code.Value = oRec.Fields.Item("Code").Value.ToString
                        dgCOLUMNS.Columns.Item("Origin").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        dgCOLUMNS.Columns.Item("Destination").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        dgCOLUMNS.Columns.Item("Service Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox

                        '--------------------------------------------------------------------------------------------
                        LoadSerType()
                        '---------------------------------------------------------------------------------------------

                        Fill_ComboBox()

                    Else

                        oStr = "SELECT Origin, Destination, SerType1 as [Service Type], SerType2 as Rate, o1st [0-5 KG], o2nd [6-49 KG], o3rd [50-249 KG], o4th [250-999 KG], o5th [1000UP] " & _
                                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                            "WHERE AppNo ='" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "' AND CompanyId = '" & oCompany.CompanyDB & "'"

                        dgCOLUMNS.DataTable.ExecuteQuery(oStr)
                        oRec.DoQuery("SELECT Code " & _
                                            "FROM [" & oCompany.CompanyDB & "]..[" & txt_AppD.Value & "] " & _
                                            "WHERE AppNo ='" & txtRM.Value & "' AND Format = '" & cmb_Format.Selected.Description & "' AND LoadType = '" & cmb_LoadT.Selected.Description & "' AND CompanyId = '" & oCompany.CompanyDB & "'")
                        txt_Code.Value = oRec.Fields.Item("Code").Value.ToString
                        dgCOLUMNS.Columns.Item("Origin").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        dgCOLUMNS.Columns.Item("Destination").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        dgCOLUMNS.Columns.Item("Service Type").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox

                        '--------------------------------------------------------------------------------------------
                        LoadSerType()
                        '---------------------------------------------------------------------------------------------

                        Fill_ComboBox()

                    End If
                End If
            End If
        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
        frmMaintenance.Freeze(False)
    End Function

    Public Function LoadSerType() As Boolean
        Try

            cmb_Format = frmMaintenance.Items.Item("cmb_Format").Specific
            cmb_LoadT = frmMaintenance.Items.Item("cmb_LoadT").Specific

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim oLoad = cmb_LoadT.Value.ToString.Substring(0, 3)
            Dim oQuery As String = "SELECT * FROM [" & oCompany.CompanyDB & "]..[@APP_SERVICETYPE] " & _
                                            "WHERE Code LIKE '" & oLoad & "%'"

            oCombo = dgCOLUMNS.Columns.Item("Service Type")

            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            oRec.DoQuery(oQuery)

            If oRec.RecordCount > 0 Then
                While oRec.EoF = False
                    oCombo.ValidValues.Add(oRec.Fields.Item(0).Value.ToString, oRec.Fields.Item(1).Value.ToString)
                    oRec.MoveNext()
                End While
            End If

            oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description

        Catch ex As Exception
            SAP_APP.SetMessage(ex)
        End Try
    End Function

#Region "For Comments"
    Private Function Comments() As Boolean
        'Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
        'Dim sCFL_ID As String = oCFLEvento.ChooseFromListUID
        'frmMaintenance = SBO_Application.Forms.Item(FormUID)
        'Dim oCFL As SAPbouiCOM.ChooseFromList = frmMaintenance.ChooseFromLists.Item(sCFL_ID)

        'If oCFLEvento.BeforeAction = False Then
        '    Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
        '    Dim val As String

        '    Try
        '        Val = oDataTable.GetValue("U_APP_RATEMATRICREF", 0)
        '    Catch ex As Exception

        '    End Try
        '    If (pVal.ItemUID = "bt_Search") Then
        '        frmMaintenance.DataSources.UserDataSources.Item("EditDS").ValueEx = val
        '    End If

        'End If
        'If CheckRM() = True Then

        'End If
    End Function
#End Region

End Module
