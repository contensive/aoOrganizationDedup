
Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationMergeTool
    '
    Public Class form2Class
        Inherits formBaseClass
        '
        '
        '
        Friend Overrides Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date, ByRef application As applicationClass) As Integer
            Dim nextFormId As Integer = srcFormId
            Try
                Dim button As String
                Dim csMain As CPCSBaseClass = cp.CSNew
                Dim csDuplicate As CPCSBaseClass = cp.CSNew
                Dim isInputOK As Boolean = True
                '
                Dim mainOrgId As Integer = cp.Visit.GetInteger("Organization Merge Main Organization ID")
                Dim duplicateOrgID As Integer = cp.Visit.GetInteger("Organization Merge Duplicate Organization ID")
                '
                Dim existSeletedIds As Boolean = False
                Dim fieldId As Integer = 0
                Dim fieldsSelected As New List(Of settingClass)
                Dim actionid As Integer = 0
                Dim sqlStr As String = ""
                '
                If Not String.IsNullOrEmpty(cp.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    fieldsSelected = deserializeFieldsSelected(cp, cp.Site.GetProperty("Organization Merge Tool  Fields Selected"))
                    existSeletedIds = True
                End If

                '
                ' ajax routines return a different name for button
                '
                button = cp.Doc.GetText("ajaxButton")
                If button = "" Then
                    button = cp.Doc.GetText(rnButton)
                End If

                If button = buttonProcess Then
                    '
                    ' Build the Sql From duplicate Organization
                    '
                    For Each onesetting As settingClass In fieldsSelected
                        If cp.Doc.GetInteger("select-" & onesetting.id.ToString) <> 0 Then
                            sqlStr &= ", " & onesetting.name
                        End If
                    Next
                    If Not String.IsNullOrEmpty(sqlStr) Then
                        sqlStr = sqlStr.Substring(2, sqlStr.Length - 2)
                    End If
                    sqlStr = "Select " & sqlStr & " from Organizations where id = " & duplicateOrgID
                    '
                    cp.Utils.AppendLog("processForm2.log", sqlStr)
                    '
                    ' generate the list of fiel with values
                    For Each onesetting As settingClass In fieldsSelected
                        'read the acion for each field
                        actionid = cp.Doc.GetInteger("select-" & onesetting.id.ToString)
                        '
                        ' 0 : Do Nothing
                        ' 1 : Replace
                        ' 2 : Replace if value is empty
                        '
                        Select Case actionid
                            Case 1
                                ' Replace
                                If csMain.Open("Organizations", "id = " & mainOrgId) Then
                                    If csDuplicate.OpenSQL(sqlStr) Then
                                        ' 
                                        csMain.SetField(onesetting.name, csDuplicate.GetText(onesetting.name))
                                        ' 
                                    End If
                                    Call csDuplicate.Close()
                                End If
                                Call csMain.Close()
                            Case 2
                                ' Replace only is not exist a value in the main organization
                                ' Replace
                                If csMain.Open("Organizations", "id = " & mainOrgId) Then
                                    If csDuplicate.OpenSQL(sqlStr) Then
                                        ' 
                                        If String.IsNullOrEmpty(csMain.GetText(onesetting.name)) Then
                                            csMain.SetField(onesetting.name, csDuplicate.GetText(onesetting.name))
                                        End If
                                        '
                                    End If
                                    Call csDuplicate.Close()
                                End If
                                Call csMain.Close()
                            Case Else
                                ' Do Nothing
                        End Select

                    Next
                    '
                    ' Inactive Duplicate Organization
                    '
                    If cp.Doc.GetBoolean("inactive") Then
                        If csDuplicate.Open("Organizations", "id =" & duplicateOrgID) Then
                            Call csDuplicate.SetField("Active", "0")
                        End If
                        Call csDuplicate.Close()
                    End If
                    '
                    ' delete Duplicate Organization
                    '
                    cp.Utils.AppendLog("form2Class.log", "before delete org")
                    If cp.Doc.GetBoolean("delete") Then
                        cp.Utils.AppendLog("form2Class.log", "delete orgid: " & duplicateOrgID)
                        'If csDuplicate.Open("Organizations", "id =" & duplicateOrgID) Then
                        '    Call csDuplicate.Delete()
                        'End If
                        'Call csDuplicate.Close()
                        cp.Content.Delete("Organizations", "id=" & duplicateOrgID)
                    End If
                    cp.Utils.AppendLog("form2Class.log", "after delete org")
                    '
                    ' Move Duplicate organization users to main Organization
                    '
                    If cp.Doc.GetBoolean("move") Then
                        If csDuplicate.Open("People", "organizationId=" & duplicateOrgID) Then
                            Do
                                '
                                Call csDuplicate.SetField("organizationID", mainOrgId.ToString)
                                '
                                Call csDuplicate.GoNext()
                            Loop While csDuplicate.OK
                        End If
                        Call csDuplicate.Close()
                    End If
                    '
                    '
                    '
                End If

                '
                ' determine the next form
                '
                Select Case button
                    Case buttonProcess
                        nextFormId = formIdThree
                    Case buttonBack
                        nextFormId = formIdOne
                End Select
            Catch ex As Exception
                errorReport(ex, cp, "processForm")
            End Try
            Return nextFormId
        End Function
        '
        '
        '
        Friend Overrides Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date, ByRef application As applicationClass) As String
            Dim returnHtml As String = ""
            Try
                Dim layout As CPBlockBaseClass = cp.BlockNew
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim body As String
                '
                Dim existSeletedIds As Boolean = False
                Dim fieldId As Integer = 0
                Dim fieldsSelected As New List(Of settingClass)

                '
                Dim htmlTableHeader As String = ""
                Dim htmlTableRow As String = ""
                Dim htmlTable As String = ""
                Dim htmlHeader As String = ""
                Dim htmlFooter As String = ""
                Dim htmlExtra As String = ""
                '
                Dim mainOrgId As Integer = cp.Visit.GetInteger("Organization Merge Main Organization ID")
                Dim duplicateOrgID As Integer = cp.Visit.GetInteger("Organization Merge Duplicate Organization ID")
                '
                Dim mainValue As String = ""
                Dim duplicateValue As String = ""
                Dim resultValue As String = ""
                '
                Dim mainName As String = ""
                Dim duplicateName As String = ""
                '
                Dim mainTotal As Integer = 0
                Dim duplicateTotal As Integer = 0
                '
                Dim sqlStr As String = ""
                '
                If Not String.IsNullOrEmpty(cp.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    fieldsSelected = deserializeFieldsSelected(cp, cp.Site.GetProperty("Organization Merge Tool  Fields Selected"))
                    existSeletedIds = True
                End If
                '
                htmlHeader = "<div class""bold""><h1>Organization Merge Tool</h1></div>" _
                    & "<div><p>Select Action to apply to the fields.</p></div>" _
                    & "<br/>"
                '
                ' Pull total of Users under organization
                '
                sqlStr = "select count(*) as total from ccMembers where OrganizationID = " & mainOrgId
                If cs.OpenSQL(sqlStr) Then
                    mainTotal = cs.GetInteger("total")
                End If
                Call cs.Close()
                '
                sqlStr = "select count(*) as total from ccMembers where OrganizationID = " & duplicateOrgID
                If cs.OpenSQL(sqlStr) Then
                    duplicateTotal = cs.GetInteger("total")
                End If
                Call cs.Close()
                '

                ' generate the list of fiel with values
                For Each onesetting As settingClass In fieldsSelected
                    '
                    If cs.Open("Organizations", "id=" & mainOrgId) Then
                        mainValue = cs.GetText(onesetting.name)
                        mainName = cs.GetText("name")
                    End If
                    Call cs.Close()
                    '
                    '
                    If cs.Open("Organizations", "id=" & duplicateOrgID) Then
                        duplicateValue = cs.GetText(onesetting.name)
                        duplicateName = cs.GetText("name")
                    End If
                    Call cs.Close()
                    '
                    resultValue = mainValue
                    If onesetting.actionId = 1 Then
                        resultValue = duplicateValue
                    ElseIf onesetting.actionId = 2 Then
                        If String.IsNullOrEmpty(mainValue) Then
                            resultValue = duplicateValue
                        Else
                            resultValue = mainValue
                        End If
                    Else
                        resultValue = mainValue
                    End If
                    '
                    htmlTableRow &= getBodyRow_5Column(onesetting.caption, "<input type=""text"" id=""js-dup-" & onesetting.id.ToString & """ value=""" & duplicateValue & """ readonly>", cp.Html.SelectList("select-" & onesetting.id.ToString, onesetting.actionId.ToString, "Use merged co. value, Use merged co. value if empty", "Keep current value", "actionSelectClass", "js-action-" & onesetting.id.ToString), "<input type=""text"" id=""js-main-" & onesetting.id.ToString & """ value=""" & mainValue & """ readonly>", "<input id=""js-text-" & onesetting.id.ToString & """ type=""=""text"" value=""" & resultValue & """ readonly>")
                Next
                htmlTableHeader = getHeaderRow_5Column("Column Name", "Organization to Merge", "Action for Field Value", "Organization to Keep", "New Field Value")
                '
                htmlTable = wrapInDivTable(htmlTableHeader, htmlTableRow)
                '

                htmlExtra = "<p> <span>" & cp.Html.CheckBox("inactive", True, "") & "</span> Deactivate duplicate organization after process.</p> <br/>" _
                        & "<p> <span>" & cp.Html.CheckBox("delete", False, "") & "</span> Delete duplicate organization after process.</p> <br/>" _
                        & "<p> Total of users in  <span class=""orgName"">" & mainName & "</span>: " & mainTotal & " <p>" _
                        & "<p> Total of users in  <span class=""orgName"">" & duplicateName & "</span>: " & duplicateTotal & " <p>" _
                        & "<p> <span>" & cp.Html.CheckBox("move", True, "") & "</span> Move users from merged organization to main organization.</p> <br/>"

                htmlFooter = "<div>" _
                        & cp.Html.Button("Back", "Back", "button", "js-backForm2") _
                        & cp.Html.Button("Process", "Process", "button", "js-processForm2") _
                        & "</div>"
                '
                body = htmlHeader & htmlTable & htmlExtra & htmlFooter
                body &= cp.Html.Hidden(rnSrcFormId, dstFormId.ToString)
                returnHtml = cp.Html.Form(body, , , "mfaForm2", rqs)
            Catch ex As Exception
                errorReport(ex, cp, "getForm")
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Function deserializeFieldsSelected(ByVal CP As CPBaseClass, ByVal value As String) As List(Of settingClass)
            Dim o As New List(Of settingClass)
            '
            Try
                '
                '   for example see testDataSerializerClass
                '
                o = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of settingClass))(value)
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in deserializeFieldsSelected")
                Catch errObj As Exception
                End Try
            End Try
            '
            Return o
        End Function
        '
        '
        '

        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.multiFormAjaxSample." & method)
        End Sub
    End Class
End Namespace
