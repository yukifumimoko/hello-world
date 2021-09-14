VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub 受信フォルダ()
     
    ' 各種オブジェクト変数
    Dim outlookApp As Outlook.Application    ' Outlookオブジェクトを入れる
    Dim myNameSpace As Outlook.Namespace     ' nameSpaceオブジェクトを入れる
    Dim myFolder As Outlook.Folder           ' フォルダを指定するのに使う
 
    ' outlookオブジェクトを作成してSetする呪文
    Set outlookApp = CreateObject("Outlook.Application")
 
    ' nameSpaceオブジェクトを取得してSetする呪文
    Set myNameSpace = outlookApp.GetNamespace("MAPI")
     
    ' フォルダーの指定(GetDefaultFolderの引数が、各フォルダに対応している)
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
     
    ' フォルダを開く(ウィンドウが表示される)
    myFolder.Display
     
    ' メモリ上に残ったオブジェクトを解放
    Set outlookApp = Nothing
    Set myNameSpace = Nothing
    Set myFolder = Nothing
     
End Sub


Sub 添付ファイル保存()

    Dim objInbox As Object
    Dim objFolder As Object
    Dim strPath As String
    Dim i As Long
     
    ' 各種オブジェクト変数
    Dim outlookApp As Outlook.Application    ' Outlookオブジェクトを入れる
    Dim myNameSpace As Outlook.Namespace     ' nameSpaceオブジェクトを入れる
    Dim myFolder As Outlook.Folder           ' フォルダを指定するのに使う
 
    ' outlookオブジェクトを作成してSetする呪文
    Set outlookApp = CreateObject("Outlook.Application")
 
    ' nameSpaceオブジェクトを取得してSetする呪文
    Set myNameSpace = outlookApp.GetNamespace("MAPI")
     
    ' フォルダーの指定(GetDefaultFolderの引数が、各フォルダに対応している)
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
     
     
    Set objInbox = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    '添付ファイルがあるメールのフォルダを指定します。2階層以上ある場合は「.Folders.Item(＜フォルダ名＞)」を追加してください。
    Set objFolder = objInbox.Folders.Item("個人")
    
    '添付ファイルの保存先をパスで指定します。
    strPath = "C:\Users\smuh1\Desktop\TestData\"
     
    For Each objItem In objFolder.Items
        For i = 1 To objItem.Attachments.Count
            '添付ファイルに拡張子がある場合のみ処理します。
            If InStr(objItem.Attachments.Item(i), ".") <> 0 Then
                objItem.Attachments.Item(i).SaveAsFile strPath & objItem.Attachments.Item(i)
            End If
        Next i
    Next objItem
 
    Set objItem = Nothing
    Set objInbox = Nothing
    Set objFolder = Nothing
 
End Sub
