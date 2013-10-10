Sub Img_in_Commentbox()

'Excel Image Comment (Scaled) Cell Script
'John McElmurray
'johnmce@microsoft.com
'Performs proper aspect ratio scaling of image (to max of 300 px, constant in script)
'Select the cell you wish to comment with an image and run the script

'Remove any old comment before adding the new one
Application.ActiveCell.ClearComments

'Pop the file picker
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False          'Only one file
    .InitialFileName = CurDir         'directory to open the window
    .Filters.Clear                    'Cancel the filter
    .Filters.Add Description:="Images", Extensions:="*.jpg,*.png", Position:=1
    .Title = "Choose image"
          
    If .Show = -1 Then TheFile = .SelectedItems(1) Else TheFile = 0
End With
    
'No file selected
If TheFile = 0 Then
    MsgBox ("No image selected")
    Exit Sub
End If

'Create the new comment
Range(ActiveCell.Address).AddComment
Range(ActiveCell.Address).Comment.Visible = False

'Get a ref to comment and add the picture file as a backgroudn
Set shp = [ActiveCell].Comment.Shape
shp.Fill.UserPicture TheFile

'Enumerate through files in folder, find the correct file, access its properties to find picture dimensions
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

strParent = FSO.GetParentFolderName(TheFile)
strArgFileName = FSO.GetFileName(TheFile)
Set objFolder = objShell.Namespace(strParent)

Dim PictureDimensions As String

For Each strFileName In objFolder.Items
    '0 is the param for Name, 1 for Dimension
    If objFolder.GetDetailsOf(strFileName, 0) = strArgFileName Then
        PictureDimensions = objFolder.GetDetailsOf(strFileName, 31)
    End If
Next

'Uncomment these lines to see all the file information in MessageBoxes
'Dim arrHeaders(39)
'For i = 0 To UBound(arrHeaders)
'    arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
'Next
'For i = 0 To UBound(arrHeaders)
'    MsgBox i & vbTab & arrHeaders(i) _
'    & ": " & objFolder.GetDetailsOf(objFolderItem, i)
'Next

'Remove "?WIDTH x HEIGHT?" special characters
Dim clean As String
clean = Mid(PictureDimensions, 2, Len(PictureDimensions) - 2)


Dim width As Integer
If IsNumeric(Split(clean, "x")(0)) Then
    width = CInt(Split(clean, "x")(0))
End If
    
Dim height As Integer
If IsNumeric(Split(clean, "x")(1)) Then
    height = CInt(Split(clean, "x")(1))
End If

'Relative resize of picture
Dim newwidth As Double
newwidth = 300
Dim aspect As Double
aspect = width / newwidth
newheight = height / aspect

' Set width and height
shp.width = newwidth
shp.height = newheight

' Lock ratio only after sizes are set proper
shp.LockAspectRatio = msoTrue

End Sub
