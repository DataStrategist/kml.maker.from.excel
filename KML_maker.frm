VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KML_maker 
   Caption         =   "kml maker 20000!!!"
   ClientHeight    =   8175
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   13440
   OleObjectBlob   =   "KML_maker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KML_maker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub userform_activate()

'###########################################################
'KML MAKER
'by Amit Kohli, FAO-AQUASTAT 2013
'based on xls2kml v.1.02 (https://groups.google.com/forum/#!topic/kml-support-third-party-applications/hlV85ESMfkA%5B151-175-false%5D)
'
'Use this as you please but:
'   1) please provide credit where it's due
'   2) there's no guarantees at all. You are responsible for how you use this file... if you corrupt any dataset it's your fault.
'   3) Tell me how you're using this, along with questions and/or comments to: mexindian@gmail.com
'###########################################################

    Me.Width = 372.6
    
    If Me.TextBox18.Value = "" Then
        Me.ToggleButton1.Value = True
        Me.ToggleButton1.Value = False
    
        Me.ComboBox1.AddItem ("Black")
        Me.ComboBox1.AddItem ("Normal")
        Me.ComboBox1.AddItem ("Blue")
        Me.ComboBox1.AddItem ("Green")
        Me.ComboBox1.AddItem ("Red")
        
        Me.ComboBox2.AddItem ("Black")
        Me.ComboBox2.AddItem ("White")
        Me.ComboBox2.AddItem ("Blue")
        Me.ComboBox2.AddItem ("Green")
        Me.ComboBox2.AddItem ("Red")
        
        Dim a
        a = Replace(ActiveWorkbook.Name, ".xls", "ƒ")
        a = Split(a, "ƒ")
        
        Me.TextBox2.Value = a(0)
        Me.TextBox13.Value = a(0)
    End If

End Sub

Private Sub ToggleButton2_Click()
If ToggleButton2 Then
    Me.Width = 675
    Me.Frame2.Visible = True
Else
    Me.Width = 372.6
    Me.Frame2.Visible = False
End If
End Sub



Private Sub CommandButton1_Click()

'OK! GO!

Dim arr(100)
Dim i, ii, num_cols, ii_ctr, st_col  As Integer
Dim n_col, lat_col, lon_col, Rest_start_col, Rest_stop_col As Integer
Dim pn, tit, pt_tit, pt_sy, pt_sc, txt_sc, pt_col, txt_col As String
Dim pt_txt_alpha
Dim fs, a As Object
Dim sp_pt_col, sp_col

Range(Me.TextBox18).Activate
pn = Me.TextBox1 & "\" & Me.TextBox13
If Right(pn, 4) = ".kml" Then pn = Left(pn, Len(pn) - 4)
tit = Me.TextBox2
pt_tit = Me.TextBox3
pt_sy = Me.TextBox4

txt_sc = Me.TextBox16
pt_sc = Me.TextBox6

txt_col = Me.ComboBox1.Value
pt_col = Me.ComboBox2.Value
pt_txt_alpha = Hex(Me.TextBox17 / 100 * 255)

'issue colours... this is the mentality in case more colors are needed:
'aabbggrr, where aa=alpha (00 to ff); bb=blue (00 to ff); gg=green (00 to ff); rr=red (00 to ff)

If txt_col = "White" Then txt_col = pt_txt_alpha & "ffffff"
If txt_col = "Black" Then txt_col = pt_txt_alpha & "000000"
If txt_col = "Blue" Then txt_col = pt_txt_alpha & "ff0000"
If txt_col = "Green" Then txt_col = pt_txt_alpha & "00ff00"
If txt_col = "Red" Then txt_col = pt_txt_alpha & "0000ff"

If pt_col = "Normal" Then pt_col = pt_txt_alpha & "ffffff"
If pt_col = "Black" Then pt_col = pt_txt_alpha & "000000"
If pt_col = "Blue" Then pt_col = pt_txt_alpha & "ff0000"
If pt_col = "Green" Then pt_col = pt_txt_alpha & "00ff00"
If pt_col = "Red" Then pt_col = pt_txt_alpha & "0000ff"


Me.ComboBox1.AddItem ("Blue")
Me.ComboBox1.AddItem ("Green")
Me.ComboBox1.AddItem ("Red")
Me.ComboBox1.AddItem ("Yellow")
st_col = ActiveCell.Column


lat_col = Range(Me.TextBox8.Value & 1).Column
lon_col = Range(Me.TextBox9.Value & 1).Column

If Me.ToggleButton1 Then
    n_col = Range(Me.TextBox7.Value & 1).Column
    If Me.TextBox10.Value <> "" Then Rest_start_col = Range(Me.TextBox10.Value & 1).Column
    If Me.TextBox11.Value <> "" Then Rest_stop_col = Range(Me.TextBox11.Value & 1).Column
End If
'pre-validate
If Me.TextBox1.Value = "" Or _
    Me.TextBox1.Value = "" Or _
    Me.TextBox1.Value = "" Or _
    Me.TextBox1.Value = "" Or _
    Me.TextBox1.Value = "" Or _
    Me.TextBox1.Value = "" Then
        MsgBox ("please fill in mandatory fields")
        Exit Sub
End If

If Application.WorksheetFunction.Count(Range(Me.TextBox8.Value & ":" & Me.TextBox8.Value)) = 0 Or Application.WorksheetFunction.Count(Range(Me.TextBox8.Value & ":" & Me.TextBox8.Value)) = 0 Then Exit Sub

On Error Resume Next
Cells.Replace What:="&", Replacement:="and", LookAt:=xlPart, MatchCase:=False
On Error GoTo 0

 If Me.TextBox10.Value <> "" Then
    ActiveCell.Offset(0, Rest_start_col - ActiveCell.Column).Activate
    
    num_cols = Rest_stop_col - Rest_start_col
    
    For ii = 0 To num_cols
        arr(ii) = ActiveCell.Offset(0, ii).Value
    Next
End If

Set fs = CreateObject("Scripting.FileSystemObject") '-----------open file
Set a = fs.CreateTextFile(pn & ".kml", True, True)

'-------------start loading

a.writeline ("<?xml version='1.0' encoding='UTF-8'?>")
a.writeline ("<kml xmlns='http://earth.google.com/kml/2.1'>")
a.writeline ("<Folder id='layer fmain'>")

a.writeline ("<name>" & tit & "</name>")


If Me.TextBox19.Value <> "" Then a.writeline ("<description>" & Me.TextBox19.Value & "</description>")

a.writeline ("<visibility>1</visibility>")
a.writeline ("<open>0</open>")

ActiveCell.Offset(1, 0).Activate

i = 1

While ActiveCell.Offset(0, st_col - ActiveCell.Column).Value <> ""  '-------------START REPETITION
    sp_pt_col = ""
    If Me.ToggleButton1 And Me.TextBox5.Value <> "" Then
        If Left(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 1) = "g" Or Left(Me.TextBox5, 1) = "G" Then
            sp_pt_col = pt_txt_alpha & "00ff00"
            sp_col = Mid(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 2, 3)
        ElseIf Left(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 1) = "b" Or Left(Me.TextBox5, 1) = "B" Then
            sp_pt_col = pt_txt_alpha & "ff0000"
            sp_col = Mid(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 2, 3)
        ElseIf Left(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 1) = "r" Or Left(Me.TextBox5, 1) = "R" Then
            sp_pt_col = pt_txt_alpha & "0000ff"
            sp_col = Mid(Range(Me.TextBox5.Value & ActiveCell.Row).Value, 2, 3)
        Else
            sp_col = Range(Me.TextBox5.Value & ActiveCell.Row).Value
        End If
    Else
        sp_col = ""
    End If
    
    
    If ActiveCell.Offset(0, st_col - ActiveCell.Column).Value <> "" And ActiveCell.Offset(0, lat_col - ActiveCell.Column).Value <> "" Then 'skips cells w/out Latitude (as a proxy for lat/long)
        a.writeline ("<Placemark id='layer p" & i & "'>")
        a.writeline ("<Style id='sn_" & i & "'>")
        a.writeline (" <IconStyle>")
        a.writeline ("  <scale>" & pt_sc & "</scale>")
        'if I specify an icon, use that, also if I specify a default direction, then use that... otherwise use whatever is in this box.
        If pt_sy <> "" Then
            If Me.ToggleButton1 Then
                If sp_col = "" Then
                    a.writeline ("  <Icon>" & pt_sy & "</Icon>")
                Else
                    a.writeline ("  <Icon>http://earth.google.com/images/kml-icons/track-directional/track-" & sp_col & ".png</Icon>")
                End If
            Else
                a.writeline ("  <Icon>" & pt_sy & "</Icon>")
            End If
        End If
        If sp_pt_col = "" Then
            a.writeline ("  <color>" & pt_col & "</color>")
        Else
            a.writeline ("  <color>" & sp_pt_col & "</color>")
        End If
        a.writeline (" </IconStyle>")
        a.writeline (" <LabelStyle>")
        If txt_col <> "" Then a.writeline ("  <color>" & txt_col & "</color>")
        a.writeline ("  <scale>" & txt_sc & "</scale>")
        a.writeline (" </LabelStyle>")
        a.writeline ("</Style>")
        
        a.writeline ("<Style id='sh_" & i & "'>")
        a.writeline (" <IconStyle>")
        a.writeline ("  <scale>" & pt_sc * 2 & "</scale>")
        If sp_pt_col = "" Then
            a.writeline ("  <color>" & pt_col & "</color>")
        Else
            a.writeline ("  <color>" & sp_pt_col & "</color>")
        End If
        'if I specify a default direction, then use that... otherwise use whatever is in this box.
        If pt_sy <> "" Then
            If Me.ToggleButton1 Then
                If sp_col = "" Then
                    a.writeline ("  <Icon>" & pt_sy & "</Icon>")
                Else
                    a.writeline ("  <Icon>http://earth.google.com/images/kml-icons/track-directional/track-" & sp_col & ".png</Icon>")
                End If
            Else
                a.writeline ("  <Icon>" & pt_sy & "</Icon>")
            End If
        End If
            
        a.writeline (" </IconStyle>")
        
        a.writeline (" <LabelStyle>")
        If txt_col <> "" Then a.writeline ("  <color>" & txt_col & "</color>")
        a.writeline ("  <scale>" & txt_sc * 2 & "</scale>")
        a.writeline (" </LabelStyle>")
        
        If Me.ToggleButton1 Then
            a.writeline (" <BalloonStyle>")
            a.writeline ("   <text><![CDATA[" & pt_tit & "<b><font color=""#0000FF"" face=""Verdana"" size=""+2""> " & _
                "$[name]</font></b> <br/><br/>  <font face=""Verdana"">$[description]</font>]]></text>")
            a.writeline (" </BalloonStyle>")
        End If
        a.writeline ("</Style>")
        
        a.writeline ("<StyleMap id='myicon_" & i & "'>")
        a.writeline (" <Pair>")
        a.writeline ("  <key>normal</key>")
        a.writeline ("  <styleUrl>#sn_" & i & "</styleUrl>")
        a.writeline (" </Pair>")
        a.writeline (" <Pair>")
        a.writeline ("  <key>highlight</key>")
        a.writeline ("  <styleUrl>#sh_" & i & "</styleUrl>")
        a.writeline (" </Pair>")
        a.writeline ("</StyleMap>")
        
         If Me.ToggleButton1 Then
            a.writeline ("<name>" & ActiveCell.Offset(0, n_col - ActiveCell.Column).Value & "</name>")
        Else
            a.writeline ("<name></name>")
        End If
        a.writeline ("<visibility>1</visibility>")
        a.writeline ("<Snippet maxLines='0' id='s" & i & "'>")
        a.writeline ("</Snippet>")
        a.writeline ("<description>")
        a.writeline ("<![CDATA[")
        If Me.TextBox10.Value <> "" Then
            For ii_ctr = 0 To ii - 1
                a.writeline ("<b>" & arr(ii_ctr) & "</b>: " & ActiveCell.Offset(0, ii_ctr) & "<br>")
            Next
        End If
        a.writeline ("]]>")
        a.writeline ("</description>")
        a.writeline ("<styleUrl>#myicon_1</styleUrl>")
        a.writeline ("<Point>")
        a.writeline ("<coordinates>")
        a.writeline (ActiveCell.Offset(0, lon_col - ActiveCell.Column).Value & "," & ActiveCell.Offset(0, lat_col - ActiveCell.Column).Value & ",0")
        a.writeline ("</coordinates>")
        a.writeline ("</Point>")
        a.writeline ("</Placemark>")
        a.writeline ("")
        a.writeline ("")
        i = i + 1
    End If
    
    ActiveCell.Offset(1, 0).Activate

Wend

        a.writeline ("</Folder>")
        a.writeline ("</kml>")
    Close
    If Me.CheckBox1.Value Then
        Unload Me
    Else
        Me.Hide
    End If
        
MsgBox ("Done!")
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub


Function GetFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function

Private Sub CommandButton3_Click()
    Me.TextBox1.Value = GetFolder(Left(ActiveWorkbook.Path, 3))
End Sub


Private Sub ToggleButton1_Click()
If Me.ToggleButton1 Then
    Me.Frame1.Visible = True
    Me.Height = 427.5
    Me.CheckBox1.Top = 384
    Me.CommandButton1.Top = 384
    Me.CommandButton2.Top = 384
Else
    Me.Frame1.Visible = False
    Me.Height = 200
    Me.CheckBox1.Top = 150
    Me.CommandButton1.Top = 150
    Me.CommandButton2.Top = 150
End If

End Sub
