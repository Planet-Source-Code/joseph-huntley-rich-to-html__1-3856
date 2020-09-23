VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Rich To HTML"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHTML 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbRichText 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.Label lblHTML 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1760
      Width           =   615
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   0
      X2              =   4200
      Y1              =   1670
      Y2              =   1670
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   4200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   0
      X2              =   4200
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   4200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblRichText 
      BackStyle       =   0  'Transparent
      Caption         =   "Rich Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*            Rich To HTML by Joseph Huntley              *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  October 4, 1999                                *
'*  Level: Beginner                                       *
'**********************************************************
'*   The form here are only used to demonstrate how to    *
'* use the function 'RichToHTML'. You may copy the        *
'* function into your project for use. If you need any    *
'* help please e-mail me.                                 *
'**********************************************************
'* Notes: None                                            *
'**********************************************************

Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String
'**********************************************************
'*            Draw Percent by Joseph Huntley              *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'**********************************************************
'*   You may use this code freely as long as credit is    *
'* given to the author, and the header remains intact.    *
'**********************************************************

'--------------------- The Arguments -----------------------
'rtbRichTextBox     - The rich textbox control to convert.
'lngStartPosition   - The character position to start from.
'lngEndPosition     - The character position to end at.
'-----------------------------------------------------------
'Returns:     The rich text converted to HTML.

'Description: Converts rich text to HTML.

Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

'check for lngStartPosition ad lngEndPosition

If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)

lngLastFontColor& = -1 'no color

   For lngCurText& = lngStartPosition& To lngEndPosition&
       rtbRichTextBox.SelStart = lngCurText&
       rtbRichTextBox.SelLength = 1
   
          If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
             intLastAlignment% = rtbRichTextBox.SelAlignment
              
                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"
                   Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"
                   Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
                End Select
                
          End If
   
          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 strHTML$ = strHTML$ & "<b>"
               Else
                 strHTML$ = strHTML$ & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 strHTML$ = strHTML$ & "<u>"
               Else
                 strHTML$ = strHTML$ & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 strHTML$ = strHTML$ & "<i>"
               Else
                 strHTML$ = strHTML$ & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 strHTML$ = strHTML$ & "<s>"
               Else
                 strHTML$ = strHTML$ & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """>"
         End If

         If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If
 
     strHTML$ = strHTML$ + rtbRichTextBox.SelText

   Next lngCurText&

RichToHTML = strHTML$

End Function
Function RGBtoHEX(lngColor As Long)



Dim strHex As String

'get hexidecimal value
strHex$ = Hex(lngColor&)

'fill in
strHex$ = String$(6 - Len(strHex$), "0") & strHex$

'swap first and third hex values.
strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
    
RGBtoHEX = strHex$

End Function



Private Sub cmdConvert_Click()
  txtHTML.Text = RichToHTML(rtbRichText, 0&, Len(rtbRichText.Text))
End Sub

Private Sub Form_Load()

  'set the text in rtbRichTextBox
  
  With rtbRichText
     .Text = "Click on the 'convert' button to convert this richtext to HTML."
     .SelStart = 0
     .SelLength = Len(.Text)
     .SelFontName = "Arial"
     .SelFontSize = 10
     .SelAlignment = rtfCenter
     .SelStart = InStr(.Text, "convert") - 1
     .SelLength = Len("convert")
     .SelFontName = "Courier New"
     .SelColor = vbBlue
     .SelStart = InStr(.Text, "HTML") - 1
     .SelLength = 4
     .SelFontName = "Courier New"
     .SelUnderline = True
     .SelStart = .SelStart + 1
     .SelLength = 1
     .SelColor = vbRed
     .SelStart = .SelStart + 1
     .SelLength = 1
     .SelColor = vbBlue
     .SelStart = .SelStart + 1
     .SelLength = 1
     .SelColor = vbGreen
     .SelStart = 0
     .SelLength = 0
  End With


End Sub
