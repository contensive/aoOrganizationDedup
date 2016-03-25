'Public Class OrganizationDedupSettingsHandler


Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationDedup
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
                ' all form classes inherit from formBase, so this one form object can be used for every form type required
                ' refresh query string is everying needed on the query string to refresh this page. Used for links, etc to always submit to the same page.
                ' rightNow is convenient when you need to test time dependant options, you can force the rightNow with a site property, etc
                ' srcFormId is the form number that submitted the request (if it was a form submission) and should be a hidden on all forms
                ' dstFormId is a way to force a form to display without a form submission
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
                ' get previously started application
                '
                'application = getApplication(CP, False)
                '
                ' if there is no application, only allow form one
                '
                'If application.id = 0 Then
                'If srcFormId <> formIdOne Then
                '    srcFormId = formSettings
                'End If
                '    dstFormId = formIdOne
                'End If

                If Not String.IsNullOrEmpty(CP.Site.GetProperty("Organization Dedup Fields Selected")) Then
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
