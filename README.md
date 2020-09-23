<div align="center">

## Secrets of the ShellExecute API


</div>

### Description

This is the most nearly complete article of the ShellExecute() function on PSC. I compiled this using various articles from MSDN. Code examples include opening default app, opening default app with error reporting, opening a defined app, and find all files dialog box.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-08-25 07:58:52
**By**             |[Michael Heusdens](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-heusdens.md)
**Level**          |Intermediate
**User Rating**    |4.9 (185 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Secrets\_of1230648272002\.zip](https://github.com/Planet-Source-Code/michael-heusdens-secrets-of-the-shellexecute-api__1-38095/archive/master.zip)





### Source Code

<h3><b>The Windows API ShellExecute() function</b></h3>
<p> </p>
<p><b>Summary</b></p>
<p><b> </b>You can use the Windows API ShellExecute() function
to start the application associated with a given document extension without
knowing the name of the associated application. The Windows API ShellExecute()
function differs from the Shell() function in that you can pass the ShellExecute()
function the name of a document, and it will start the associated application,
and then pass the filename to the application. You can also specify the working
folder for the application.</p>
<p><b>1. Declaring the ShellExecute function</b></p>
<p><b> </b>The ShellExecute() function has 6 paramenters, many
of which are optional:</p>
<div align="right">
 <table border="1" width="695">
 <tr>
 <td width="11%" align="center"><i><u>Parameter</u></i></td>
 <td width="89%" align="center"><i><u>Description</u></i></td>
 </tr>
 <tr>
 <td width="11%" valign="top">hwnd</td>
 <td width="89%">The Window handle to a parent. This window recieves any
 message boxes an application produces (for example, for error
 reporting).
 <br><br>
 </td>
 </tr>
 <tr>
 <td width="11%" valign="top">lpOperation </td>
 <td width="89%">Specifies the action (verb) to perform. With the exception
 of NULL, these are passed as strings:
 <div align="right">
 <table border="0" width="96%">
 <tr>
 <td width="9%" valign="top">Open </td>
 <td width="91%">Opens the file specified by the lpFile parameter.
 The file can be an executable file, a document file, or a
 folder.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Print</td>
 <td width="91%">Prints the document file specified by lpFile. If
 lpFile is an executable file, it opens the file, as if
 "Open" had been specified.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Explore </td>
 <td width="91%">Explores the folder specified by lpFile.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Find</td>
 <td width="91%">Initiates a search starting from the folder
 specified by lpFile.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Edit</td>
 <td width="91%">Launches an editor and opens the document
 specified by lpFile for editing.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Play</td>
 <td width="91%">For methods supporting a play function, such as
 sound files.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">0&</td>
 <td width="91%">lpOperation can also be NULL (0& if declared
 As Any, or vbNullString if declared As String). In these cases
 the call performs the default verb action, which is usually
 "Open". The default action can be seen by viewing
 specific extension in Explorer's Tool / Folder Options / File
 Types.<br><br></td>
 </tr>
 <tr>
 <td width="9%" valign="top">Properties </td>
 <td width="91%">Displays the properties of the file specified by
 lpFile.</td>
 </tr>
 </table>
 </div><br><br>
 </td>
 </tr>
 <tr>
 <td width="11%" valign="top">lpFile</td>
 <td width="89%">A string that specifies an executable file, a document
 file, or a folder.
 <br><br>
 </td>
 </tr>
 <tr>
 <td width="11%" valign="top">lpParameters </td>
 <td width="89%">If the lpFile parameter specifies an executable file,
 lpParameters is a string that specifies the parameters to be passed to
 the application (the startup parameters and/or a document file). If
 lpFile specifies a document file, lpParameters should be NULL (0&).
 <br><br>
 </td>
 </tr>
 <tr>
 <td width="11%" valign="top">lpDirectory</td>
 <td width="89%">A string that specifies the default (working) directory.
 <br><br>
 </td>
 </tr>
 <tr>
 <td width="11%" valign="top">nShowCmd</td>
 <td width="89%">Flags that specify how the application is to be shown when
 it is opened. If this parameter is not specified, the application uses
 its default value.</td>
 </tr>
 </table>
</div>
<p> Now lets declare the ShellExecute() function:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">Declare Function ShellExecute Lib "shell32.dll" Alias _<br>
 "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _<br>
 As String, ByVal lpFile As String, ByVal lpParameters _<br>
 As String, ByVal lpDirectory As String, ByVal nShowCmd _<br>
 As Long) As Long</td>
 </tr>
 </table>
</div>
<p> </p>
<p><b>2. Setting Constants for nShowCmd's Flags </b></p>
<p> The flags for the nShowCmd parameter are as follows:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="21%" align="center"><i><u>Flag</u></i></td>
 <td width="79%" align="center"><i><u>Description</u></i></td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_HIDE</td>
 <td width="79%">Hides the window and activates another window.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWNORMAL</td>
 <td width="79%">Activates and displays the window. If the window is
 minimized or maximized, Windows restores it to its original size and
 position. An application should specify this flag when displaying the
 window for the first time.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWMINIMIZED</td>
 <td width="79%">Activates the window and displays it as a minimized
 window.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWMAXIMIZED </td>
 <td width="79%">Activates the window and displays it as a maximized
 window.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWNOACTIVE</td>
 <td width="79%">Displays the window at its most recent size and position.
 The active window remains active.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOW</td>
 <td width="79%">Activates the window and displays it at its current size
 and position.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_MINIMIZE</td>
 <td width="79%">Minimizes the specified window and activates the next
 top-level window in the z-order.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWMINNOACTIVE </td>
 <td width="79%">Displays the window as a minimized window. The active
 window remains active.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWNA</td>
 <td width="79%">Displays the window in its current state. The active
 window remains active.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_RESTORE</td>
 <td width="79%">Activates and displays the window. If the window is
 minimized or maximized, Windows restores it to its original size and
 position. An application should specify this flag when restoring a
 minimized window.
 <p> </p>
 </td>
 </tr>
 <tr>
 <td width="21%" valign="top">SW_SHOWDEFAULT</td>
 <td width="79%">Sets the show state based on the SW_ flag specified in the
 STARTUPINFO stucture passed to the CreateProcess function by the program
 that started the application.</td>
 </tr>
 </table>
</div>
<p> The flags are numerical values. There are two ways of
using these flags for the nShowCmd parameter. To make it easier for updating
your application, we'll set these values to flag names. Place the following in
your code:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">Const SW_HIDE = 0<br>
 Const SW_SHOWNORMAL = 1<br>
 Const SW_SHOWMINIMIZED = 2<br>
 Const SW_SHOWMAXIMIZED = 3<br>
 Const SW_SHOWNOACTIVE = 4<br>
 Const SW_SHOW = 5<br>
 Const SW_MINIMIZE = 6<br>
 Const SW_SHOWMINNOACTIVE = 7<br>
 Const SW_SHOWNA = 8<br>
 Const SW_RESTORE = 9<br>
 Const SW_SHOWDEFAULT = 10
 </td>
 </tr>
 </table>
</div>
<p> </p>
<p><b>3. Return Values</b></p>
<p> The return values are received by the hwnd parameter.
Return values return a value greater than 32 if successful, or an error value
that is less than or equal to 32 otherwise. The following table lists the error
values:</p>
<div align="right">
 <table border="1" width="95%">
 <tr>
 <td width="28%" valign="top">0</td>
 <td width="72%">The operating system is out of memory or resources.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">ERROR_FILE_NOT_FOUND</td>
 <td width="72%">The specified file was not found.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">ERROR_PATH_NOT_FOUND </td>
 <td width="72%">The specified path was not found.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">ERROR_BAD_FORMAT</td>
 <td width="72%">The .exe file is invalid (non-Win32 .exe or error in .exe
 image).
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_ACCESSDENIED</td>
 <td width="72%">The operating system denied access to the specified file.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_ASSOCINCOMPLETE </td>
 <td width="72%">The file name association is incomplete or invalid.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_DDEBUSY</td>
 <td width="72%">The DDE transaction could not be completed because other
 DDE transactions were being processed.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_DDEFAIL</td>
 <td width="72%">The DDE transaction failed.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_DDETIMEOUT</td>
 <td width="72%">The DDE transaction could not be completed because the
 request timed out.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_DLLNOTFOUND</td>
 <td width="72%">The specified dynamic-link library was not found.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_FNF</td>
 <td width="72%">The specified file was not found.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_NOASSOC</td>
 <td width="72%">There is no application associated with the given file
 name extension. This error will also be returned if you attempt to print
 a file that is not printable.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_OOM</td>
 <td width="72%">There was not enough memory to complete the operation.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_PNF</td>
 <td width="72%">The specified path was not found.
 <br><br></td>
 </tr>
 <tr>
 <td width="28%" valign="top">SE_ERR_SHARE</td>
 <td width="72%">A sharing violation occurred.</td>
 </tr>
 </table>
</div>
<p> </p>
<p><b>4. </b><b>Code Examples</b></p>
<p><b> Open Default Application Code</b></p>
<p> 1. Start a new Standard EXE project.</p>
<p> 2. Add a Command Button to the form.</p>
<p> 3. Copy the following code to the Code window of Form1:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">Private Declare Function ShellExecute Lib
 "shell32.dll" _<br>
 Alias "ShellExecuteA" _<br>
 (ByVal hwnd As Long, _<br>
 ByVal lpOperation As String, _<br>
 ByVal lpFile As String, _<br>
 ByVal lpParameters As String, _<br>
 ByVal lpDirectory As String, _<br>
 ByVal nShowCmd As Long) As Long
 <p>Private Const SW_HIDE = 0<br>
 Private
 Const SW_SHOWNORMAL = 1<br>
 Private
 Const SW_SHOWMINIMIZED = 2<br>
 Private
 Const SW_SHOWMAXIMIZED = 3<br>
 Private
 Const SW_SHOWNOACTIVE = 4<br>
 Private
 Const SW_SHOW = 5<br>
 Private
 Const SW_MINIMIZE = 6<br>
 Private
 Const SW_SHOWMINNOACTIVE = 7<br>
 Private
 Const SW_SHOWNA = 8<br>
 Private
 Const SW_RESTORE = 9</p>
 <p>Private Sub Command1_Click()<br>
 Dim opbrowser As Long<br>
 'Run default browser<br>
 opbrowser = ShellExecute(Me.hwnd, "open", _<br>
 "http://www.yahoo.com", "", "C:\", SW_SHOWNORMAL)<br>
 End Sub</p>
 </td>
 </tr>
 </table>
</div>
<p><b> Open Default Application Code with Return Values (MSDN
Example)</b></p>
<p> 1. Start a new Standard EXE project.</p>
<p> 2. Copy the following code to the Code window of Form1:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">Option Explicit
 <p>Private Declare Function ShellExecute Lib
 "shell32.dll" _<br>
 Alias "ShellExecuteA" _<br>
 (ByVal hwnd As Long, _<br>
 ByVal lpOperation As String, _<br>
 ByVal lpFile As String, _<br>
 ByVal lpParameters As String, _<br>
 ByVal lpDirectory As String, _<br>
 ByVal nShowCmd As Long) As Long</p>
 <p>Private Declare Function GetDesktopWindow Lib "user32" ()
 As Long</p>
 <p>Private Const SW_HIDE = 0<br>
 Private
 Const SW_SHOWNORMAL = 1<br>
 Private
 Const SW_SHOWMINIMIZED = 2<br>
 Private
 Const SW_SHOWMAXIMIZED = 3<br>
 Private
 Const SW_SHOWNOACTIVE = 4<br>
 Private
 Const SW_SHOW = 5<br>
 Private
 Const SW_MINIMIZE = 6<br>
 Private
 Const SW_SHOWMINNOACTIVE = 7<br>
 Private
 Const SW_SHOWNA = 8<br>
 Private
 Const SW_RESTORE = 9</p>
 <p>Const SE_ERR_FNF = 2&<br>
 Const SE_ERR_PNF = 3&<br>
 Const SE_ERR_ACCESSDENIED = 5&<br>
 Const SE_ERR_OOM = 8&<br>
 Const SE_ERR_DLLNOTFOUND = 32&<br>
 Const SE_ERR_SHARE = 26&<br>
 Const SE_ERR_ASSOCINCOMPLETE = 27&<br>
 Const SE_ERR_DDETIMEOUT = 28&<br>
 Const SE_ERR_DDEFAIL = 29&<br>
 Const SE_ERR_DDEBUSY = 30&<br>
 Const SE_ERR_NOASSOC = 31&<br>
 Const ERROR_BAD_FORMAT = 11&</p>
 <p>Function StartDoc(DocName As String) As Long<br>
 Dim Scr_hDC As Long<br>
 Scr_hDC = GetDesktopWindow()<br>
 StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _<br>
 "", "C:\", SW_SHOWNORMAL)<br>
 End Function<br>
 <br>
 Private Sub Form_Click()<br>
 Dim r As Long, msg As String<br>
 r = StartDoc("C:\WINDOWS\ARCADE.BMP")<br>
 If r <= 32 Then<br>
 'There was an error<br>
 Select Case r<br>
 Case SE_ERR_FNF<br>
 msg = "File not found"<br>
 Case SE_ERR_PNF<br>
 msg = "Path not found"<br>
 Case SE_ERR_ACCESSDENIED<br>
 msg = "Access denied"<br>
 Case SE_ERR_OOM<br>
 msg = "Out of memory"<br>
 Case SE_ERR_DLLNOTFOUND<br>
 msg = "DLL not found"<br>
 Case SE_ERR_SHARE<br>
 msg = "A sharing violation occurred"<br>
 Case SE_ERR_ASSOCINCOMPLETE<br>
 msg = "Incomplete or invalid file association"<br>
 Case SE_ERR_DDETIMEOUT<br>
 msg = "DDE Time out"<br>
 Case SE_ERR_DDEFAIL<br>
 msg = "DDE transaction failed"<br>
 Case SE_ERR_DDEBUSY<br>
 msg = "DDE busy"<br>
 Case SE_ERR_NOASSOC<br>
 msg = "No association for file extension"<br>
 Case ERROR_BAD_FORMAT<br>
 msg = "Invalid EXE file or error in EXE image"<br>
 Case Else<br>
 msg = "Unknown error"<br>
 End Select<br>
 MsgBox msg<br>
 End If<br>
 End Sub</p>
 </td>
 </tr>
 </table>
</div>
<p><b> Open Defined Application Code</b></p>
<p> 1. Start a new Standard EXE project.</p>
<p> 2. Add a Command Button to the form.</p>
<p> 3. Copy the following code to the Code window of Form1:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">Private Declare Function ShellExecute Lib
 "shell32.dll" _<br>
 Alias "ShellExecuteA" _<br>
 (ByVal hwnd As Long, _<br>
 ByVal lpOperation As String, _<br>
 ByVal lpFile As String, _<br>
 ByVal lpParameters As String, _<br>
 ByVal lpDirectory As String, _<br>
 ByVal nShowCmd As Long) As Long
 <p>Private Const SW_HIDE = 0<br>
 Private Const SW_SHOWNORMAL = 1<br>
 Private Const SW_SHOWMINIMIZED = 2<br>
 Private Const SW_SHOWMAXIMIZED = 3<br>
 Private Const SW_SHOWNOACTIVE = 4<br>
 Private Const SW_SHOW = 5<br>
 Private Const SW_MINIMIZE = 6<br>
 Private Const SW_SHOWMINNOACTIVE = 7<br>
 Private Const SW_SHOWNA = 8<br>
 Private Const SW_RESTORE = 9</p>
 <p>Private Sub Command1_Click()<br>
 Dim r As Long<br>
 r = ShellExecute(Me.hwnd, "open", _<br>
 "C:\Program Files\Internet Explorer\iexplore.exe", _<br>
 "C:\readme.txt", "C:\",
 SW_SHOWNORMAL)<br>
 End Sub</p>
 </td>
 </tr>
 </table>
</div>
<p><b> Find All Files Dialog Box Code (MSDN Example)</b></p>
<p> 1. Start a new Standard EXE project.</p>
<p> 2. Add a Command Button, a Drive ListBox, and a Directory
ListBox to the form.</p>
<p> 3. Copy the following code to the Code window of Form1:</p>
<div align="right">
 <table border="0" width="95%">
 <tr>
 <td width="100%" bgcolor="#C0C0C0">
 <p align="left">Private Declare Function ShellExecute Lib
 "shell32.dll" _<br>
 Alias "ShellExecuteA" _<br>
 (ByVal hwnd As Long, _<br>
 ByVal lpOperation As String, _<br>
 ByVal lpFile As String, _<br>
 ByVal lpParameters As String, _<br>
 ByVal lpDirectory As String, _<br>
 ByVal nShowCmd As Long) As Long</p>
 <p align="left">Private Const SW_SHOWNORMAL = 1<br>
 Private Const SW_SHOWMINIMIZED = 2<br>
 Private Const SW_SHOWMAXIMIZED = 3<br>
 Private Const SW_SHOW = 5<br>
 Private Const SW_MINIMIZE = 6<br>
 Private Const SW_SHOWMINNOACTIVE = 7<br>
 Private Const SW_SHOWNA = 8<br>
 Private Const SW_RESTORE = 9<br>
 Private Const SW_SHOWDEFAULT = 10</p>
 <p align="left">Private Sub Command1_Click()<br>
 Call ShellExecute(Me.hwnd, _<br>
 "find", _<br>
 Dir1.Path, _<br>
 vbNullString, _<br>
 vbNullString, _<br>
 SW_SHOWNORMAL)<br>
 End Sub</p>
 <p align="left">Private Sub Drive1_Change()<br>
 Dir1.Path = Drive1.Drive<br>
 End Sub</p>
 <p align="left">Private Sub Form_Load()<br>
 Command1.Caption = "Search
 Selected Path"<br>
 End Sub</p></td>
 </tr>
 </table>
</div>
<p> </p>

