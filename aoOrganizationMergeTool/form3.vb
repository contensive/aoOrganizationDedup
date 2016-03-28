
Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationMergeTool
    '
    Public Class form3Class
        Inherits formBaseClass
        '
        '
        '
        Friend Overrides Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date, ByRef application As applicationClass) As Integer
            Dim nextFormId As Integer = srcFormId
            Try
                Dim button As String
                '
                button = cp.Doc.GetText("ajaxButton")
                If button = "" Then
                    button = cp.Doc.GetText(rnButton)
                End If

                '
                ' determine the next form
                '
                If button = buttonOK Then
                    ' Reset main and duplicate org Id
                    '
                    cp.Visit.SetProperty("Organization Merge Main Organization ID", "")
                    cp.Visit.SetProperty("Organization Merge Duplicate Organization ID", "")
                    '
                End If

                Select Case button
                    Case buttonOK
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
                '
                Dim mainValue As String = ""
                '
                Dim mainName As String = ""
                '
                Dim mainTotal As Integer = 0
                '
                Dim sqlStr As String = ""
                '
                If Not String.IsNullOrEmpty(cp.Site.GetProperty("Organization Merge Tool  Fields Selected")) Then
                    fieldsSelected = deserializeFieldsSelected(cp, cp.Site.GetProperty("Organization Merge Tool  Fields Selected"))
                    existSeletedIds = True
                End If
                '
                ' Pull total of Users under organization
                '
                sqlStr = "select count(*) as total from ccMembers where OrganizationID = " & mainOrgId
                If cs.OpenSQL(sqlStr) Then
                    mainTotal = cs.GetInteger("total")
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
                    htmlTableRow &= getBodyRow_3Column("#", onesetting.caption, "<input type=""text"" id=""js-main-" & onesetting.id.ToString & """ value=""" & mainValue & """ readonly>")
                Next
                htmlTableHeader = getHeaderRow_3Column("#", "Column Name", "Main Organization Value")
                '
                htmlTable = wrapInDivTable(htmlTableHeader, htmlTableRow)
                '
                '
                htmlHeader = "<div class""bold""><h1>Organization Merge Tool Result</h1></div>" _
                    & "<br/>" _
                    & "<p> Actual field values for <span class=""bold"">" & mainName & "</span>.<p>"
                '
                htmlExtra = "<p> Total of users in <span class=""orgName"">" & mainName & "</span>: " & mainTotal & " <p>"

                htmlFooter = "<div>" _
                        & cp.Html.Button("OK", "OK", "button", "js-okForm3") _
                        & "</div>"
                '
                body = htmlHeader & htmlTable & htmlExtra & htmlFooter
                body &= cp.Html.Hidden(rnSrcFormId, dstFormId.ToString)
                returnHtml = cp.Html.Form(body, , , "mfaForm3", rqs)
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
