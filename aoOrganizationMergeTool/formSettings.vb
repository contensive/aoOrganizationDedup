Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationMergeTool
    '
    Public Class settingFormClass
        Inherits formBaseClass
        '
        '
        '
        Friend Overrides Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightnow As Date, ByRef application As applicationClass) As Integer
            Dim nextFormId As Integer = srcFormId
            Try
                Dim button As String
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim isInputOK As Boolean = True
                Dim num As Integer = 0
                Dim fieldsSelected As New List(Of settingClass)
                Dim fieldSelected As settingClass
                '
                ' ajax routines return a different name for button
                '
                button = cp.Doc.GetText("ajaxButton")
                If button = "" Then
                    button = cp.Doc.GetText(rnButton)
                End If
                '
                If button = buttonSave Then
                    ' load values
                    If cs.Open("Content Fields", "ContentID=" & cp.Content.GetRecordID("Content", "Organizations") & " and Authorable=1 ") Then
                        Do
                            '
                            num += 1
                            If cp.Doc.GetInteger("check-" & cs.GetInteger("id")) = 1 Then
                                fieldSelected = New settingClass
                                fieldSelected.id = cs.GetInteger("id")
                                fieldSelected.name = cs.GetText("name")
                                fieldSelected.caption = cs.GetText("caption")
                                fieldSelected.actionId = cp.Doc.GetInteger("select-" & cs.GetInteger("id"))
                                fieldsSelected.Add(fieldSelected)
                            End If
                            Call cs.GoNext()
                        Loop While cs.OK
                    End If
                    Call cs.Close()
                    ' save values
                    If fieldsSelected.Count > 0 Then
                        cp.Site.SetProperty("Organization Merge Tool  Fields Selected", serializeFieldsSelected(cp, fieldsSelected))
                    Else
                        cp.Site.SetProperty("Organization Merge Tool  Fields Selected", "")
                    End If
                End If

                If isInputOK Then
                    '
                    '
                    '
                    ' determine the next form
                    '
                    Select Case button
                        Case buttonSave
                            nextFormId = formIdOne
                        Case buttonBack
                            nextFormId = formIdOne
                    End Select
                End If
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
            Dim htmlRow As String = ""
            Dim htmlTable As String = ""
            Dim htmlHeader As String = ""
            Dim htmlFooter As String = ""
            Try
                Dim layout As CPBlockBaseClass = cp.BlockNew
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim body As String
                Dim num As Integer = 0

                '
                Dim existSeletedIds As Boolean = False
                Dim idSelected As Boolean = False
                Dim actionID As Integer = 0
                Dim fieldsSelected As New List(Of settingClass)
                '
                If Not String.IsNullOrEmpty(cp.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    fieldsSelected = deserializeFieldsSelected(cp, cp.Site.GetProperty("Organization Merge Tool  Fields Selected"))
                    existSeletedIds = True
                End If
                '
                htmlHeader = "<div class""bold""><h1>Organization Merge Tool Settings</h1></div>" _
                    & "<div><p>Select Organization fields to be part of the merge process.</p></div>" _
                    & "<br/>"
                '
                If cs.Open("Content Fields", "ContentID=" & cp.Content.GetRecordID("Content", "Organizations") & " and Authorable=1 ") Then
                    Do
                        '
                        num += 1
                        actionID = 0
                        If existSeletedIds Then
                            idSelected = False
                            For i = 0 To (fieldsSelected.Count - 1)
                                If fieldsSelected(i).id = cs.GetInteger("id") Then
                                    idSelected = True
                                    actionID = fieldsSelected(i).actionId
                                    Exit For
                                End If
                            Next

                            htmlRow &= getBodyRow_4Column(num.ToString,
                                                          cp.Html.CheckBox("check-" & cs.GetInteger("id"), idSelected, , "js-check-" & cs.GetInteger("id")),
                                                          cs.GetText("caption"),
                                                          cp.Html.SelectList("select-" & cs.GetInteger("id"), actionID.ToString, "Use merged co. value, Use merged co. value if empty", "Keep current value", "", ""))
                        Else
                            htmlRow &= getBodyRow_4Column(num.ToString,
                                                          cp.Html.CheckBox("check-" & cs.GetInteger("id"), False, , "js-check-" & cs.GetInteger("id")),
                                                          cs.GetText("caption"),
                                                          cp.Html.SelectList("select_" & cs.GetInteger("id"), "", "Use merged co. value, Use merged co. value if empty", "Keep current value", "", ""))
                        End If
                        '
                        Call cs.GoNext()
                    Loop While cs.OK
                End If
                Call cs.Close()
                '
                htmlTable = wrapInDivTable(getHeaderRow_4Column("#", "Select", "Field Name", "Default Action"), htmlRow)
                '
                htmlFooter = "<div>" _
                    & cp.Html.Button("Save", "Save", "button", "js-saveSetting") _
                    & "</div>"

                '
                body = htmlHeader & htmlTable & htmlFooter
                body &= cp.Html.Hidden(rnSrcFormId, dstFormId.ToString)
                returnHtml = cp.Html.Form(body, , , "mfaSetting", rqs)
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
        Private Function serializeFieldsSelected(ByVal CP As CPBaseClass, ByVal package As List(Of settingClass)) As String
            Dim s As String = ""
            '
            Try
                s = Newtonsoft.Json.JsonConvert.SerializeObject(package).Replace(vbCrLf, "")
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in serializeFieldsSelected")
                Catch errObj As Exception
                End Try
            End Try
            '
            Return s
        End Function

        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.multiFormAjaxSample." & method)
        End Sub
    End Class
End Namespace

