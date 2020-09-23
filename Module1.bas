Attribute VB_Name = "Module1"
'Option Explicit
'
'Modules: Automating Internet Explorer 5
'      Author (s)
'      Dev Ashish
'
'
'      Microsoft Internet Explorer comes with a fairly comprehensive, although sparsely documented, Object Model.  If you've used the Web Browser control in Access, you are already familiar with the capabilities of IE's Object Model.  All of the functionality in IE's object model (not counting external support, like scripting support etc.) is provided by the following two dlls:
'
'      shdocvw.dll  (Microsoft Internet Controls)
'      mshtml.tlb (Microsoft HTML Object Library)
'      You can automate IE to save a HTML file locally (read the comments in the code), inspect all the elements, and parse out a particular item at runtime.
'
'      Here 's some sample code that loops through all the IE windows currently open looking for a browser that has the string URL_TO_SEARCH in it's address bar. If it finds such a window, it prompts the user to save the HTML as a local file. Additionally, it will go through all the HTML elements in that page and try to find an anchor whose description is ANCHOR_DESC_TO_SEARCH. If it finds this element, it will print out the URL the anchor is pointing to in the debug window.
'
'      (Also look at the article API: Read-Set Internet Explorer URL from code for an API approach under a somewhat similar scenario.)
'
'' ****** Code Start *********
'' This code was originally written by Dev Ashish
'' It is not to be altered or distributed,
'' except as part of an application.
'' You are free to use it in any application,
'' provided the copyright notice is left unchanged.
''
'' Code Courtesy of
'' Dev Ashish
''
'Sub sTestIEAutomation()
'      ' Requires two references
'      ' shdocvw.dll - Microsoft Internet Controls
'      ' mshtml.tlb  - Microsoft HTML Object Library
'      '
'      On Error GoTo ErrHandler
'Dim objShellWins As SHDocVw.ShellWindows
'Dim objIE As SHDocVw.InternetExplorer
'Dim objDoc As Object
'Dim i As Integer
'Dim strOut As String
'Dim intFree As Integer
'Dim clsDialog As CDialog  ' Wrapper around GetOpen/SaveFileName
'Const URL_TO_SEARCH = "http://www.mvps.org/access"
'Const ANCHOR_DESC_TO_SEARCH = "Comprehensive Links"
'
'      ' Instantiate
'      Set objShellWins = New SHDocVw.ShellWindows
'      ' There might be multiple IE windows open
'      For Each objIE In objShellWins
'         With objIE
'            ' Try to locate the browser with a specific address
'            ' in it's AddressBar. You can also Navigate to a new
'            ' address
'            If (InStr(1, _
'            .LocationURL, _
'            URL_TO_SEARCH, vbTextCompare)) Then
'
'               ' Get a reference to the HTMLDocument contained within
'               ' the InternetExplorer instance
'               Set objDoc = .document
'               If (TypeOf objDoc Is HTMLDocument) Then
'                  ' Limitations of running the following command:
'                  ' Call objIE.ExecWB( _
'                  OLECMDID_SAVEAS, _
'                  OLECMDEXECOPT_PROMPTUSER)
'                  ' IE's "SaveAs" dialog doesn't allow you to
'                  ' retrieve the filename the user typed in
'                  ' so use our own code for the SaveAs dialog
'                  ' The CDialog class is simply a wrapper around
'                  ' the code listed at the following URL
'                  ' http://www.mvps.org/access/api/api0001.htm
'                  '
'                  Set clsDialog = New CDialog
'                  With clsDialog
'                     .hWnd = hWndAccessApp
'                     .StartDir = CurDir
'                     .ModeOpen = False
'                     .DefaultExtension = "htm"
'                     .Title = "Please select a folder to save the file"
'                     .Filter = "HTML Files (*.htm, *.html)|*.htm"
'                     strOut = .Action
'                  End With
'                  If Len(strOut) Then
'                     ' Now that we have a filename,
'                     ' Save out the HTML as a persisted file
'                     intFree = FreeFile
'                     Open strOut For Output As #intFree
'                     Write #intFree, objDoc.body.parentElement.innerHTML
'                     Close #intFree
'                     ' Alternatively, you could also just
'                     ' inpect the HTM at runtime via the property
'                     With objDoc.All
'                        For i = 1 To .length
'                           If (TypeOf .Item(i) Is HTMLAnchorElement) Then
'                              If .Item(i).nodeName = "A" Then
'                                 ' Only look for a link which has the description
'                                 ' "Comprehensive Links" attached to it
'                                 If (InStr(1, _
'                                 .Item(i).innerText, _
'                                 ANCHOR_DESC_TO_SEARCH, _
'                                 vbTextCompare)) Then
'                                    ' Print out the URL
'                                    Debug.Print objDoc.All.Item(i).href
'                                    ' Bail out
'                                    Exit For
'                                 End If
'                              End If
'                           End If
'                        Next
'                     End With
'                  End If
'               End If
'               Exit For
'            End If
'         End With
'      Next
'
'ExitHere:
'      On Error Resume Next
'      Close #intFree
'      Set clsDialog = Nothing
'      Set objDoc = Nothing
'      Set objIE = Nothing
'      Set objShellWins = Nothing
'      Exit Sub
'ErrHandler:
'      With Err
'         MsgBox "Error: " & .Number & vbCrLf & .Description, _
'         vbCritical Or vbOKOnly, .Source
'      End With
'      Resume ExitHere
'End Sub
'' ****** Code End *********
'
'
'
