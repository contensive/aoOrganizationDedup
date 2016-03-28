
Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationMergeTool
    '
    Public Class form1Class
        Inherits formBaseClass
        '
        '
        '
        Friend Overrides Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightnow As Date, ByRef application As applicationClass) As Integer
            Dim nextFormId As Integer = srcFormId
            Try
                Dim button As String
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim isInputOK As Boolean = False
                Dim mainOrgId As Integer = cp.Doc.GetInteger("mainOrgID")
                Dim duplicateOrgID As Integer = cp.Doc.GetInteger("duplicateOrgID")
                '
                ' ajax routines return a different name for button
                '
                button = cp.Doc.GetText("ajaxButton")
                If button = "" Then
                    button = cp.Doc.GetText(rnButton)
                End If
                '
                If (mainOrgId <> duplicateOrgID) And (mainOrgId <> 0) And (duplicateOrgID <> 0) Then
                    cp.Visit.SetProperty("Organization Merge Main Organization ID", mainOrgId.ToString)
                    cp.Visit.SetProperty("Organization Merge Duplicate Organization ID", duplicateOrgID.ToString)
                    isInputOK = True
                End If
                '
                '
                ' determine the next form
                '
                Select Case button
                    Case buttonNext
                        '
                        If isInputOK Then
                            nextFormId = formIdTwo
                        Else
                            nextFormId = formIdOne
                        End If
                        '
                    Case buttonSettings
                        nextFormId = formSettings
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
                Dim htmlRow1 As String = ""
                Dim htmlRow2 As String = ""
                Dim htmlTable As String = ""
                Dim htmlHeader As String = ""
                Dim htmlFooter As String = ""
                Dim mainOrgId As Integer = cp.Visit.GetInteger("Organization Merge Main Organization ID")
                Dim duplicateOrgID As Integer = cp.Visit.GetInteger("Organization Merge Duplicate Organization ID")

                htmlHeader = "<div class""bold""><h1>Organization Merge Tool</h1></div>" _
                            & "<div><p>Select Main Organization and duplicate organization.</p></div>" _
                            & "<br/>"
                '
                htmlRow1 = getBodyRow_3Column("", "Main Organization: ", cp.Html.SelectContent("mainOrgID", mainOrgId.ToString, "Organizations", "", "Select One", , "js-mainOrgID"))
                htmlRow2 = getBodyRow_3Column("", "Duplicate Organization:", cp.Html.SelectContent("duplicateOrgID", duplicateOrgID.ToString, "Organizations", "", "Select One", , "js-duplicateOrgID"))
                '
                htmlTable = wrapInDivTable(htmlRow1, htmlRow2)
                '
                htmlFooter = "<div>" _
                        & cp.Html.Button("Settings", "Settings", "button", "js-settingsForm1") _
                        & cp.Html.Button("Next", "Next", "button", "js-nextForm1") _
                        & "</div>"
                '
                body = htmlHeader & htmlTable & htmlFooter
                body &= cp.Html.Hidden(rnSrcFormId, dstFormId.ToString)
                returnHtml = cp.Html.Form(body, , , "mfaForm1", rqs)
            Catch ex As Exception
                errorReport(ex, cp, "getForm")
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Private Sub errorReport(ByVal ex As Exception, ByVal cp As CPBaseClass, ByVal method As String)
            cp.Site.ErrorReport(ex, "error in aoManagerTemplate.multiFormAjaxSample." & method)
        End Sub
    End Class
End Namespace
