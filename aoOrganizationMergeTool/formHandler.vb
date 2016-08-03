Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationMergeTool
    '
    Public Class formHandlerClass
        Inherits AddonBaseClass
        '
        ' Ajax Handler - Remote Method
        '   returns content of inner classes to the contains originally created around them
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                '
                '
                '
                Dim body As String = ""
                Dim form As formBaseClass
                '
                Dim rqs As String = CP.Doc.GetProperty("multiformAjaxVbFrameRqs")
                '
                Dim rightNow As Date = getRightNow(CP)
                '
                Dim srcFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnSrcFormId))
                Dim dstFormId As Integer = CP.Utils.EncodeInteger(CP.Doc.GetProperty(rnDstFormId))
                '
                Dim formHandler As formHandlerClass = New formHandlerClass
                Dim application As New applicationClass
                '
                Dim cs As CPCSBaseClass = CP.CSNew
                '
                If Not String.IsNullOrEmpty(CP.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    dstFormId = formIdOne
                Else

                    ' dstFormId = formSettings
                    ' Change we need to remove setting page, snow we need show thr full list of fields

                    ' build the list and go to form one
                    '
                    Dim fieldsSelected As New List(Of settingClass)
                    Dim fieldSelected As settingClass
                    If cs.Open("Content Fields", "ContentID=" & CP.Content.GetRecordID("Content", "Organizations") & " and Authorable=1 ") Then
                        Do
                            '
                            fieldSelected = New settingClass
                            fieldSelected.id = cs.GetInteger("id")
                            fieldSelected.name = cs.GetText("name")
                            fieldSelected.caption = cs.GetText("caption")
                            fieldSelected.actionId = CP.Doc.GetInteger("select-" & cs.GetInteger("id"))
                            fieldsSelected.Add(fieldSelected)
                            Call cs.GoNext()
                        Loop While cs.OK
                    End If
                    Call cs.Close()
                    ' save values
                    If fieldsSelected.Count > 0 Then
                        CP.Site.SetProperty("Organization Merge Tool  Fields Selected", serializeFieldsSelected(CP, fieldsSelected))
                    Else
                        CP.Site.SetProperty("Organization Merge Tool  Fields Selected", "")
                    End If

                    ' after create the fields go to form 1
                    dstFormId = formIdOne
                End If
                '
                ' process forms
                '
                If (srcFormId <> 0) Then
                    Select Case srcFormId
                        Case formIdThree
                            '
                            form = New form3Class
                            dstFormId = form.processForm(CP, srcFormId, rqs, rightNow, application)
                            '
                        Case formIdTwo
                            '
                            form = New form2Class
                            dstFormId = form.processForm(CP, srcFormId, rqs, rightNow, application)
                            '
                        Case formIdOne
                            '
                            form = New form1Class
                            dstFormId = form.processForm(CP, srcFormId, rqs, rightNow, application)
                            '
                        Case Else
                            form = New settingFormClass
                            dstFormId = form.processForm(CP, srcFormId, rqs, rightNow, application)
                    End Select
                End If
                '
                ' get the next form that should appear on the page. 
                ' put the default form as the else case - to display if nothing else is selected
                '
                Select Case dstFormId
                    Case formIdThree
                        '
                        '
                        '
                        form = New form3Class
                        body = form.getForm(CP, dstFormId, rqs, rightNow, application)
                    Case formIdTwo
                        '
                        '
                        '
                        form = New form2Class
                        body = form.getForm(CP, dstFormId, rqs, rightNow, application)
                    Case formIdOne
                        '
                        '
                        '
                        form = New form1Class
                        body = form.getForm(CP, dstFormId, rqs, rightNow, application)
                    Case Else
                        '
                        ' default form
                        '
                        dstFormId = formSettings
                        form = New settingFormClass
                        body = form.getForm(CP, dstFormId, rqs, rightNow, application)
                End Select
                ''Call saveApplication(CP, application, rightNow)
                '
                ' assemble body
                '
                returnHtml = body
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in formHandlerClass.execute")
            End Try
            Return returnHtml
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
    End Class
End Namespace
