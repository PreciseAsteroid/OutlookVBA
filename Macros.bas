Attribute VB_Name = "Macros"
Sub Createlinkfromclipboard()
    Dim MyDataObject As DataObject
    Set MyDataObject = New DataObject
    MyDataObject.GetFromClipBoard
    MyDataObject.GetText
    
    
    
    Dim ObjDoc As Word.Document
    Dim ObjSel As Word.Selection
    
    If Application.ActiveInspector.EditorType = olEditorWord Then
    
        Set ObjDoc = Application.ActiveInspector.WordEditor
        Set ObjSel = ObjDoc.Windows(1).Selection
        'check selection
        'Debug.Print ObjSel.Text
        ObjDoc.Hyperlinks.Add ObjSel.Range, MyDataObject.GetText, "", "", "link", ""
    End If
    
    Set ObjDoc = Nothing
    Set ObjSel = Nothing
   
    
End Sub
