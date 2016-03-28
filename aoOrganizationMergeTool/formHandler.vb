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
                If Not String.IsNullOrEmpty(CP.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    dstFormId = formIdOne
                Else
                    dstFormId = formSettings
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
    End Class
End Namespace
