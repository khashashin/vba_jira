Option Explicit

Public bearbeiter As String
Public stichwort As String
Public issueType As String
Public projektId As Integer


Public Sub SendJiraTask()
    '*********************************
    '*          Requarments          *
    '*********************************
    'Visual Basic For Application
    'Microsoft Outlook 16.0 Object Library
    'OLE Automation
    'Microsoft Office 16.0 Object Library
    'Microsoft XML, v6.0
    'Microsoft Scripting Runtime
    'Microsoft Script Control 1.0
    'Microsoft Forms 2.0 Object Library
    
    Dim OlMail As Outlook.MailItem
    Dim OlApp As Outlook.Application
    Dim OlSel As Outlook.Selection
    Dim mySubject As String
    Dim myBody As String
    Dim OlCount As Integer

    Set OlApp = CreateObject("Outlook.Application")
    Set OlSel = OlApp.ActiveExplorer.Selection
    ChooserForm.Show
    
    For OlCount = 1 To OlSel.Count
        Set OlMail = OlSel.Item(OlCount)
    Next OlCount
    
    mySubject = CreateSubjBody(OlMail.Subject)
    myBody = CreateSubjBody(OlMail.Body)
    
    Dim JIRA As String
    JIRA = createJSON( _
                    projektId, _
                    mySubject, _
                    myBody, _
                    bearbeiter, _
                    issueType, _
                    stichwort)
                    
    CreateJiraTask "yourusername", "yourjirapasword", JIRA '<= Jira username und password
    
End Sub


Private Sub CreateJiraTask(user As String, Password As String, JSON As String)
    Dim webhook As Object
    Dim URL As String
    'Open this link to view the generated request https://webhook.site/#/55759d1a-7892-4c20-8d15-3b8b7f1bf3b3/e5e428a6-24a3-43c6-a15b-49edb1764468/0
    URL = "https://jira.com/rest/api/2/issue/"
    Set webhook = CreateObject("MSXML2.XMLHTTP.6.0")
    webhook.Open "POST", URL, False
    webhook.setRequestHeader "Content-Type", "application/json"
    webhook.setRequestHeader "User-Agent", "Outlook"
    webhook.setRequestHeader "Authorization", "Basic " + Base64Encode(user + ":" + Password)
    webhook.Send JSON
End Sub
Private Function createJSON(customProjektID As Integer, customSubject As String, customBody As String, customUser As String, customIssue As String, customStichwort As String) As String
    Dim JSON(17) As String, JSONText As String
    JSON(0) = "{"
    JSON(1) = "  ""fields"": {"
    JSON(2) = "    ""project"": {"
    JSON(3) = "      ""id"": @customProjektID"
    JSON(4) = "    },"
    JSON(5) = "    ""summary"": ""@customSubject"","
    JSON(6) = "    ""description"": ""@customBody"","
    JSON(7) = "    ""issuetype"": {"
    JSON(8) = "      ""name"": ""@issuName"""
    JSON(9) = "    },"
    JSON(10) = "    ""assignee"": {"
    JSON(11) = "      ""name"": ""@recipAlias"""
    JSON(12) = "    },"
    JSON(13) = "    ""labels"": ["
    JSON(14) = "      ""@jiraLabel"""
    JSON(15) = "    ]"
    JSON(16) = "  }"
    JSON(17) = "}"

    JSONText = Join(JSON, vbNewLine)
    JSONText = Replace(JSONText, "@customProjektID", customProjektID)
    JSONText = Replace(JSONText, "@customSubject", customSubject)
    JSONText = Replace(JSONText, "@customBody", customBody)
    JSONText = Replace(JSONText, "@issuName", customIssue)
    JSONText = Replace(JSONText, "@recipAlias", customUser)
    JSONText = Replace(JSONText, "@jiraLabel", customStichwort)

    createJSON = JSONText
End Function

Private Function CreateSubjBody(Text As String) As String
    Text = Replace(Text, """", "'")
    Text = Replace(Text, vbCr & vbLf, "\n")
    Text = Replace(Text, vbCr, "\n")
    Text = Replace(Text, vbLf, "\n")
    CreateSubjBody = Text
End Function


' The below is taken from http://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript

Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.Text
    Set oNode = Nothing
    Set oXML = Nothing
End Function



'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data

Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function



'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string

Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
