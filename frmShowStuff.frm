VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShowStuff 
   Caption         =   "Show Stuff"
   ClientHeight    =   2784
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5664
   LinkTopic       =   "Form1"
   ScaleHeight     =   2784
   ScaleWidth      =   5664
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowOpen 
      Caption         =   "Show &Open"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdShowSave 
      Caption         =   "Show &Save"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton cmdShowColor 
      Caption         =   "Show &Color"
      Height          =   372
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdShowFont 
      Caption         =   "Show &Font"
      Height          =   372
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton cmdShowPrinter 
      Caption         =   "Show &Printer"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdShowHelp 
      Caption         =   "Show &Help"
      Height          =   372
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog dlgB 
      Left            =   5040
      Top             =   600
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmShowStuff.frx":0000
      Top             =   1440
      Width           =   5532
   End
   Begin VB.Image img8Ball 
      Height          =   384
      Left            =   4080
      Picture         =   "frmShowStuff.frx":0008
      ToolTipText     =   "Exit"
      Top             =   960
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmShowStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Louis Boldt
' Common Dialog Template
' 6/7/2000
 
Dim mintFileIdIn As Integer
Dim mintFileIdOut As Integer
Dim mstrFileNameIn As String
Dim mstrFileNameOut As String

Private Sub cmdShowOpen_Click()
' ----------------------------------------------------------------------
' Flags
' Open and save Dialog Box flags use 'or' to select multiple flags
'cdlOFNAllowMultiselect   '&H200 Enable multiple selection
'cdlOFNCreatePrompt       '&H2000 Prompt if file does not exist
'cdlOFNExplorer           '&H80000 Use explorer type open dialog
'cdlOFNExtensionDifferent ' &H400 Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property.
'cdlOFNFileMustExist '&H1000 Specifies that the user can enter only names of existing files in the File Name text box
'cdlOFNHelpButton    '&H10 Causes the dialog box to display the Help button.
'cdlOFNHideReadOnly  '&H4 Hides the Read Onlycheck box.
'cdlOFNLongNames     '&H200000 Use long filenames.
'cdlOFNNoReadOnlyReturn   '&H8000  Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
'cdlOFNNoChangeDir        '&H8 Forces the dialog box to set the current directory to what it was when the dialog box was opened.
'cdlOFNNoDereferenceLinks '&H100000 Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
'cdlOFNNoLongNames     '&H40000 Do not use long file names.
'cdlOFNOverwritePrompt '&H2 Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
'cdlOFNPathMustExist   '&H800 Specifies that the user can enter only validpaths. If this flag is set and the user enters an invalid path, a warning message is displayed
'cdlOFNReadOnly        '&H1 Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
'cdlOFNShareAware      '&H4000 Specifies that sharing violation errors will be ignored.

   On Error GoTo Handle_Cancel
   With dlgB
      .DialogTitle = "What file do you want to use for Input ???"
      .Flags = cdlOFNHideReadOnly _
            Or cdlOFNFileMustExist
      .DefaultExt = "txt"  ' Default ext of file to save if none given
         
      .InitDir = "C:\DoBeDo\"
      .FileName = "DaDaDaaDoWop.txt"
      
      .Filter = "Access DB (*.mdb) |*.mdb" _
              & "|Text Files (*.txt)|*.txt" _
              & "|Dlimited (*.csv)|*.csv;*.txt;*.dat" _
              & "|All Files (*.*)|*.*"
              
      .FilterIndex = 2 'set default to  (*.txt)
      .CancelError = True
      .ShowOpen 'Show the File Open dialog
   End With
   txtNote.Text = dlgB.FileTitle & vbCrLf 'Return name of file (no path) Read only

   txtNote.Text = txtNote.Text & dlgB.FileName & vbCrLf 'full name and path
   
  
   If dlgB.Flags And cdlOFNExtensionDifferent Then 'ck extention. Same as default?
      txtNote.Text = txtNote.Text & " ext diffrent from default"
   Else
      txtNote.Text = txtNote.Text & " ext same as default"
   End If
    mstrFileNameIn = dlgB.FileName   'store the filename
'   mintFileIdIn = FreeFile
'   Open dlgB.Filename For Input As #mintFileIdIn
   
Handle_cancel_exit:
Exit Sub
   
Handle_Cancel:
   If Err = cdlCancel Then
      txtNote.Text = "Open Canceled"
      Resume Handle_cancel_exit
   Else
      MsgBox "Unexpected Error Returned" & vbCr _
           & "Nmbr " & Err.Number & vbCrLf _
           & "Desc " & Err.Description & vbCr _
           & "Srce " & Err.Source & vbCr _
           & "Time " & Now & vbCr _
           & "Path " & App.Path, _
           16, "You have a Problem!"
      
   End If
End Sub

Private Sub cmdShowSave_Click()
' ----------------------------------------------------------------------
   Dim strMod As String
     
   On Error GoTo Handle_Cancel
   With dlgB
     
      .Flags = cdlOFNOverwritePrompt  '&H2 Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.

'      Open and save Dialog Box Properties
      .CancelError = True ' Generate error on cancel
      .DefaultExt = "csv"  ' Default ext of file to save if none given
                            ' or extention given is not registered
      .DialogTitle = "Where do You want to go Today"
      
      .InitDir = "c:\DaCode\"
      .FileName = "FileNxxx"
      .MaxFileSize = 256  ' max filename size 256 default (may want to increase if multi select allowed)
      
      
      .Filter = "Dlimited (*.csv)|*.csv;*.txt;*.dat" _
             & "|Text Files (*.txt)|*.txt" _
             & "|All Files (*.*)|*.*"
      .FilterIndex = 3 'set default to  (*.)
      .CancelError = True 'Return error if cancel pressed
      .ShowSave 'Show the File Save dialog
   End With
   txtNote.Text = dlgB.FileTitle & vbCrLf 'Return name of file (no path) Read only

   txtNote.Text = txtNote.Text & dlgB.FileName & vbCrLf _
         & dlgB.Flags
   If dlgB.Flags And cdlOFNExtensionDifferent Then
      txtNote.Text = txtNote.Text & " ext diffrent from default"
   Else
      txtNote.Text = txtNote.Text & " ext same as default"
   End If
   
   
'   mintFileIdOut = FreeFile
'   If chkAppend.Value = Checked Then
'      Open dlgB.Filename For Append As #mintFileIdOut
'
'   Else
'      Open dlgB.Filename For Output As #mintFileIdOut
'
'   End If
 
   mstrFileNameOut = dlgB.FileName   'store the filename
Handle_cancel_exit:
Exit Sub
   
Handle_Cancel:
   If Err = cdlCancel Then
     txtNote.Text = "Save Canceled"
      Resume Handle_cancel_exit
   Else
      MsgBox "Unexpected Error Returned" & vbCr _
           & "Nmbr " & Err.Number & vbCrLf _
           & "Desc " & Err.Description & vbCr _
           & "Srce " & Err.Source & vbCr _
           & "Time " & Now & vbCr _
           & "Path " & App.Path, _
           16, "You have a Problem!"
      
   End If

End Sub

Private Sub cmdShowColor_Click()
' ----------------------------------------------------------------------
'   Colors
'Constant Value Description
'vbBlack   &H0      Black
'vbRed     &HFF     Red
'vbGreen   &HFF00   Green
'vbYellow  &HFFFF   Yellow
'vbBlue    &HFF0000 Blue
'vbMagenta &HFF00FF Magenta
'vbCyan    &HFFFF00 Cyan
'vbWhite   &HFFFFFF White

'Color Dialog Box flags use 'or' to select multiple flags
'cdlCCRGBInit '&H1 make .color the default rgb value
'cdlCCFullOpen '&H2 Entire dialog box is displayed, including the Define Custom Colors section
'cdlCCHelpButton '&H8 Causes the dialog box to display a Help button
'cdlCCPreventFullOpen ' &H4 Disables the Define Custom Colors command button and prevents the user from defining custom colors
   On Error GoTo Handle_Cancel
   With dlgB
      .CancelError = True
      .Flags = cdlCCRGBInit '&H1 make .color the default rgb value
      .Color = txtNote.BackColor 'default color is what's already in text box
      .ShowColor            'Display color box
    End With
   txtNote.BackColor = dlgB.Color 'set the new color
Handle_cancel_exit:
Exit Sub
   
Handle_Cancel:
   If Err = cdlCancel Then
      txtNote.Text = "Color Canceled"
      Resume Handle_cancel_exit
   Else
      MsgBox "Unexpected Error Returned" & vbCr _
           & "Nmbr " & Err.Number & vbCrLf _
           & "Desc " & Err.Description & vbCr _
           & "Srce " & Err.Source & vbCr _
           & "Time " & Now & vbCr _
           & "Path " & App.Path, _
           16, "You have a Problem!"
      
   End If
   
End Sub

Private Sub cmdShowFont_Click()
' ----------------------------------------------------------------------
'Font Dialog Box flags use 'or' to select multiple flags
'cdlCFANSIOnly ' Only Windows Character Set
'cdlCFApply    ' &H200 Enables the Apply button on the dialog box.
'cdlCFBoth     ' &H3 Causes the dialog box to list the available printer and screen fonts. The hDC property identifies thedevice context associated with the printer.
'cdlCFEffects  ' &H100 Specifies that the dialog box enables strikethrough, underline, and color effects.
'cdlCFFixedPitchOnly ' &H4000 Specifies that the dialog box selects only fixed-pitch fonts.
'cdlCFForceFontExist ' &H10000 Specifies that an error message box is displayed if the user attempts to select a font or style that doesn't exist.
'cdlCFHelpButton    ' &H4 Causes the dialog box to display a Help button.
'cdlCFLimitSize     ' &H2000 Specifies that the dialog box selects only font sizes within the range specified by the Min and Max properties.
'cdlCFNoFaceSel     ' &H80000 No font name selected.
'cdlCFNoSimulations ' &H1000 Specifies that the dialog box doesn't allow graphic device interface (GDI) font simulations.
'cdlCFNoSizeSel     ' &H200000 No font size selected.
'cdlCFNoStyleSel    ' &H100000
'cdlCFNoVectorFonts ' &H800 Specifies that the dialog box doesn't allow vector-font selections.
'cdlCFPrinterFonts  ' &H2 Causes the dialog box to list only the fonts supported by the printer, specified by the hDC property.
'cdlCFScalableOnly  ' &H20000 Specifies that the dialog box allows only the selection of fonts that can be scaled.
'cdlCFScreenFonts   ' &H1 Causes the dialog box to list only the screen fonts supported by the system.
'cdlCFTTOnly        ' &H40000 Specifies that the dialog box allows only the selection of TrueType fonts.
'cdlCFWYSIWYG       ' &H8000 Specifies that the dialog box allows only the selection of fonts that are available on both the printer and on screen.
                    ' If this flag is set, the cdlCFBoth and cdlCFScalableOnly flags should also be set.

   On Error GoTo Handle_Cancel
   With dlgB
      .CancelError = True
      
      .Flags = cdlCFBoth _
            Or cdlCFEffects     ' both printer and screen+ special effects
      
      .FontBold = txtNote.FontBold            'set the dialog box's default
      .FontItalic = txtNote.FontItalic        'values to what's already in
      .FontUnderline = txtNote.FontUnderline  'the text box
      .FontStrikethru = txtNote.FontStrikethru
      .Color = txtNote.ForeColor
      .FontName = txtNote.FontName
      .FontSize = txtNote.FontSize
      
      .Max = 24    ' maximum font size (set LimitSize first)
      .Min = 8     ' minimum font size (set LimitSize first)
      .ShowFont        'Display the Font dialog box
   End With
   txtNote.FontBold = dlgB.FontBold  'and set the text box's properties
   txtNote.FontItalic = dlgB.FontItalic
   txtNote.FontStrikethru = dlgB.FontStrikethru
   txtNote.FontUnderline = dlgB.FontUnderline
   txtNote.ForeColor = dlgB.Color
   txtNote.FontName = dlgB.FontName
   txtNote.FontSize = dlgB.FontSize
Handle_cancel_exit:
Exit Sub
   
Handle_Cancel:
   If Err = cdlCancel Then
      txtNote.Text = "Font Canceled"
      Resume Handle_cancel_exit
   Else
      MsgBox "Unexpected Error Returned" & vbCr _
           & "Nmbr " & Err.Number & vbCrLf _
           & "Desc " & Err.Description & vbCr _
           & "Srce " & Err.Source & vbCr _
           & "Time " & Now & vbCr _
           & "Path " & App.Path, _
           16, "You have a Problem!"
      
   End If

End Sub

Private Sub cmdShowPrinter_Click()
' ----------------------------------------------------------------------
'  Printer Dialog Box Flags
'Constant Value Description
'cdlPDAllPages           &H0 Returns or sets the state of the All Pagesoption button.
'cdlPDCollate            &H10 Returns or sets the state of the Collatecheck box.
'cdlPDDisablePrintToFile &H80000 Disables the Print To File check box.
'cdlPDHelpButton         &H800 Causes the dialog box to display the Help button.
'cdlPDHidePrintToFile    &H100000 Hides the Print To File check box.
'cdlPDNoPageNums         &H8 Disables the Pages option button and the associated edit control.
'cdlPDNoSelection        &H4 Disables the Selection option button.
'cdlPDNoWarning          &H80 Prevents a warning message from being displayed when there is no default printer.
'cdlPDPageNums           &H2 Returns or sets the state of the Pages option button.
'cdlPDPrintSetup         &H40 Causes the system to display the Print Setup dialog box rather than the Print dialog box.
'cdlPDPrintToFile        &H20 Returns or sets the state of the Print To File check box.
'cdlPDReturnDC           &H100 Returns adevice context for the printer selection made in the dialog box. The device context is returned in the dialog box's hDC property.
'cdlPDReturnDefault      &H400 Returns default printer name.
'cdlPDReturnIC           &H200 Returns an information context for the printer selection made in the dialog box. An information context provides a fast way to get information about the device without creating a device context. The information context is returned in the dialog box's hDC property.
'cdlPDSelection          &H1 Returns or sets the state of the Selection option button. If neither cdlPDPageNums nor cdlPDSelection is specified, the All option button is in the selected state.
'cdlPDUseDevModeCopies   &H40000 If a printer driver doesn't support multiple copies, setting this flag disables the copies edit control. If a driver does support multiple copies, setting this flag indicates that the dialog box stores the requested number of copies in the Copies property.
 
   On Error GoTo Handle_Cancel
   With dlgB
      .Copies = 1
      
      .Flags = cdlPDAllPages
      .FromPage = 1
      .ToPage = 1
      .Max = 9 ' max print range
      .Min = 1 ' min print range
      
      .CancelError = True
      .PrinterDefault = True
      
      .ShowPrinter          'Display Print dialog box
   End With
'      .PrinterDefault 'Changes / sets default printer (see VB documentation)
'      .hDC ' handle to selected printer must be used
'           ' in api call if default printer is not chosed

'      if defaut printer chosen then you can print directly to the printer object      Printer.Print txtNote.Text
Handle_cancel_exit:
Exit Sub
   
Handle_Cancel:
   If Err = cdlCancel Then
      txtNote.Text = "Print Canceled"
      Resume Handle_cancel_exit
   Else
      MsgBox "Unexpected Error Returned" & vbCr _
           & "Nmbr " & Err.Number & vbCrLf _
           & "Desc " & Err.Description & vbCr _
           & "Srce " & Err.Source & vbCr _
           & "Time " & Now & vbCr _
           & "Path " & App.Path, _
           16, "You have a Problem!"
      
   End If

End Sub

Private Sub cmdShowHelp_Click()
' ----------------------------------------------------------------------
   img8Ball.Visible = True
   txtNote.Text = "HELP YourSelf"
End Sub

Private Sub img8Ball_Click()
   End
End Sub
