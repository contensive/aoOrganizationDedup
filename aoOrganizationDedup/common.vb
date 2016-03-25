Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoOrganizationDedup


    ' This class holds the current application, so the Db only has to be read once and saved once
    ' for some requrements this can be larger and describe an entire environment, to include values from many records.
    ' for some projects, you may decide that every page does not need all the values stored in this class. In that case you can create
    ' methods in the class that load on demand. For instance, in the constructor just pass cp and exit. When firstname is used, check if the 
    ' application id is 0. If so, read all the values from the application record into the object before returning firstName.
    '
    Public Class applicationClass
        Public completed As Boolean
        Public changed As Boolean
    End Class
    '
    ' base class for the form classes. Having this means in he form handler you can only create one object, not one per form
    '
    Public MustInherit Class formBaseClass
        Friend MustOverride Function processForm(ByVal cp As CPBaseClass, ByVal srcFormId As Integer, ByVal rqs As String, ByVal rightNow As Date, ByRef application As applicationClass) As Integer

        Friend MustOverride Function getForm(ByVal cp As CPBaseClass, ByVal dstFormId As Integer, ByVal rqs As String, ByVal rightNow As Date, ByRef application As applicationClass) As String
    End Class
    '
    '
    Public Class settingClass
        Public id As Integer
        Public name As String
        Public caption As String
        Public actionId As Integer
    End Class
    '
    '
    '
    Module common
        '
        ' Button Messages
        '
        Public Const buttonOK As String = "OK"
        Public Const buttonSave As String = "Save"
        Public Const buttonCancel As String = "Cancel"
        Public Const buttonNext As String = "Next"
        Public Const buttonPrevious As String = "Previous"
        Public Const buttonContinue As String = "Continue"
        Public Const buttonRestart As String = "Restart"
        Public Const buttonFinish As String = "Finish"
        Public Const buttonBack As String = "Back"
        Public Const buttonProcess As String = "Process"
        Public Const buttonSettings As String = "Settings"
        '
        ' request names 
        '
        Public Const rnUserId As String = "userId"
        Public Const rnSrcFormId As String = "srcFormId"
        Public Const rnDstFormId As String = "dstFormId"
        Public Const rnButton As String = "button"
        '
        ' Forms
        '
        Public Const formIdOne As Integer = 1
        Public Const formIdTwo As Integer = 2
        Public Const formIdThree As Integer = 3
        Public Const formIdFour As Integer = 4
        '
        Public Const formSettings As Integer = 10
        '
        ' Layoutd
        '
        Public Const form01Layout As String = "MultiFormAjaxSample - Form 1"
        Public Const form02Layout As String = "MultiFormAjaxSample - Form 1"
        '
        '
        '
        Friend Function buffDate(ByVal sourceDate As Date) As String
            Dim returnValue As String
            '
            If sourceDate < #1/1/1900# Then
                returnValue = ""
            Else
                returnValue = sourceDate.ToShortDateString
            End If
            Return returnValue

        End Function
        '
        '
        '
        Friend Function getRightNow(ByVal cp As Contensive.BaseClasses.CPBaseClass) As Date
            Dim returnValue As Date = Date.Now()
            Try
                Dim testString As String = cp.Site.GetProperty("Sample Manager Test Mode Date", "")
                '
                ' change 'sample' to the name of this collection
                '
                If testString <> "" Then
                    returnValue = encodeMinDate(cp.Utils.EncodeDate(testString))
                    If returnValue = Date.MinValue Then
                        returnValue = Date.Now()
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "Error in getRightNow")
            End Try
            Return returnValue
        End Function
        '
        '
        '
        Friend Function encodeMinDate(ByVal sourceDate As Date) As Date
            Dim returnValue As Date = sourceDate
            If returnValue < #1/1/1900# Then
                returnValue = Date.MinValue
            End If
            Return returnValue
        End Function
        '
        '
        '
        '
        Friend Sub appendLog(ByVal cp As CPBaseClass, ByVal logMessage As String)
            Dim nowDate As Date = Date.Now.Date()
            Dim logFilename As String = nowDate.Year & nowDate.Month.ToString("D2") & nowDate.Day.ToString("D2") & ".log"
            Call cp.File.CreateFolder(cp.Site.PhysicalInstallPath & "\logs\multiformAjaxSample")
            Call cp.Utils.AppendLog("multiformAjaxSample\" & logFilename, logMessage)
        End Sub
        '
        '
        '

        Friend Function getApplication(ByVal cp As CPBaseClass, ByVal createRecordIfMissing As Boolean) As applicationClass
            Dim application As applicationClass = New applicationClass
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim csSrc As CPCSBaseClass = cp.CSNew()
                '
                ' get id of this user's application
                ' use visit property if they keep the same application for the visit
                ' use visitor property if each time they open thier browser, they get the previous application
                ' use user property if they only get to the application when they are associated to the current user (they should be authenticated first)
                '
                'application.completed = False
                'application.changed = False
                'application.id = cp.Visit.GetInteger("multiformAjaxSample ApplicationId")
                ''application.id = cp.Visitor.GetInteger("multiformAjaxSample ApplicationId")
                ''application.id = cp.user.GetInteger("multiformAjaxSample ApplicationId")
                'If (application.id <> 0) Then
                '    If Not cs.Open("MultiFormAjax Application", "(dateCompleted is null)") Then
                '        application.id = 0
                '    End If
                'End If
                'If cs.OK Then
                '    '
                '    ' application record exists, use application data
                '    '
                '    application.firstName = cs.GetText("firstName")
                '    application.lastName = cs.GetText("lastName")
                '    application.email = cs.GetText("email")
                'Else
                '    '
                '    ' application does not exist, load defaults
                '    ' this is a good place for an improvement - these could load on demain, loading only when they are needed
                '    '
                '    If csSrc.Open("people", "id=" & cp.User.Id) Then
                '        application.firstName = csSrc.GetText("firstName")
                '        application.lastName = csSrc.GetText("lastName")
                '        application.email = csSrc.GetText("email")
                '    End If
                '    Call csSrc.Close()
                'End If
                'Call cs.Close()

            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "Error in getApplication")
            End Try
            Return application
        End Function
        '
        '
        '
        Public Sub saveApplication(ByVal cp As CPBaseClass, ByVal application As applicationClass, ByVal rightNow As Date)
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                'If application.changed Then
                '    If application.id > 0 Then
                '        Call cs.Open(cnMultiFormAjaxApplications, "(id=" & application.id & ")")
                '    End If
                '    If Not cs.OK Then
                '        If cs.Insert("MultiFormAjax Application") Then
                '            '
                '            ' create a new record. 
                '            ' Set application Id incase needed later
                '            ' Set visit property to save the application Id
                '            '
                '            'application.id = cs.GetInteger("id")
                '            'Call cp.Visit.SetProperty("multiformAjaxSample ApplicationId", application.id.ToString())
                '            'Call cp.Visitor.SetProperty("multiformAjaxSample ApplicationId", application.id.ToString())
                '            'Call cp.User.SetProperty("multiformAjaxSample ApplicationId", application.id.ToString())
                '            Call cs.SetField("visitId", cp.Visit.Id.ToString())
                '        End If
                '    End If
                '    If cs.OK Then
                '        Call cs.SetField("firstName", application.firstName)
                '        Call cs.SetField("lastName", application.lastName)
                '        Call cs.SetField("email", application.email)
                '        If application.completed Then
                '            Call cs.SetField("datecompleted", rightNow.ToString)
                '        End If
                '    End If
                '    Call cs.Close()
                'End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "Error in getApplication")
            End Try
        End Sub
        '
        '
        Public Function getHeaderRow_3Column(value1 As String, value2 As String, value3 As String) As String
            Dim o As String = ""
            '
            o = " <tr><th>" & value1 & "</th><th>" & value2 & "</th><th>" & value3 & "</th></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        Public Function getBodyRow_3Column(value1 As String, value2 As String, value3 As String) As String
            Dim o As String = ""
            '
            o = " <tr><td>" & value1 & "</td><td>" & value2 & "</td><td>" & value3 & "</td></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        Public Function getHeaderRow_4Column(value1 As String, value2 As String, value3 As String, value4 As String) As String
            Dim o As String = ""
            '
            o = " <tr><th>" & value1 & "</th><th>" & value2 & "</th><th>" & value3 & "</th><th>" & value4 & "</th></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        Public Function getBodyRow_4Column(value1 As String, value2 As String, value3 As String, value4 As String) As String
            Dim o As String = ""
            '
            o = " <tr><td>" & value1 & "</td><td>" & value2 & "</td><td>" & value3 & "</td><td>" & value4 & "</td></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        Public Function getHeaderRow_5Column(value1 As String, value2 As String, value3 As String, value4 As String, value5 As String) As String
            Dim o As String = ""
            '
            o = " <tr><th>" & value1 & "</th><th>" & value2 & "</th><th>" & value3 & "</th><th>" & value4 & "</th><th>" & value5 & "</th></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        Public Function getBodyRow_5Column(value1 As String, value2 As String, value3 As String, value4 As String, value5 As String) As String
            Dim o As String = ""
            '
            o = " <tr><td>" & value1 & "</td><td>" & value2 & "</td><td>" & value3 & "</td><td>" & value4 & "</td><td>" & value5 & "</td></tr> " & vbCrLf
            '
            Return o
        End Function
        '
        '
        '
        '
        Public Function wrapInDivTable(header As String, body As String, Optional className As String = "") As String
            Dim o As String = ""
            '
            If String.IsNullOrEmpty(className) Then
                o = " <table> " & header & body & " </table> " & vbCrLf
            Else
                o = " <table class=""" & className & """> " & header & body & " </table> " & vbCrLf
            End If
            '
            Return o
        End Function
        '
        '
        Public Function wrapInDiv(htmlInner As String, cssClass As String) As String
            Dim o As String = ""
            '
            If String.IsNullOrEmpty(cssClass) Then
                o = "<div>" & htmlInner & "</div>"
            Else
                o = "<div class=""" & cssClass & """>" & htmlInner & "</div>"
            End If
            '
            Return o
        End Function
        '
        '

    End Module
End Namespace