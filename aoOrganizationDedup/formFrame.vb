Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationDedup
    '
    '
    '
    Public Class formFrameClass
        Inherits AddonBaseClass
        '
        '
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Try
                '
                ' refresh query string is everying needed on the query string to refresh this page. Used for links, etc to always submit to the same page.
                '
                Dim body As String = ""
                Dim rqs As String = CP.Doc.RefreshQueryString
                Dim formHandler As formHandlerClass = New formHandlerClass
                '
                ' use the formHandler. It is built to handle ajax calls, so setup the call like an ajax call would be, setting the Refresh Query String
                '
                Call CP.Doc.SetProperty("multiformAjaxVbFrameRqs", rqs)
                body = formHandler.Execute(CP)
                '
                ' adding multiformAjaxVbFrameRqs to the page's javascript environment is important so ajax requests back from the page can know what the frame's RQS was for links, etc.
                '
                Call CP.Doc.AddHeadJavascript("var multiformAjaxVbFrameRqs='" & rqs & "';")
                '
                returnHtml = CP.Html.div(body, , "box", "multiFormAjaxFrame")
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Error"
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace
