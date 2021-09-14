VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ��M�t�H���_()
     
    ' �e��I�u�W�F�N�g�ϐ�
    Dim outlookApp As Outlook.Application    ' Outlook�I�u�W�F�N�g������
    Dim myNameSpace As Outlook.Namespace     ' nameSpace�I�u�W�F�N�g������
    Dim myFolder As Outlook.Folder           ' �t�H���_���w�肷��̂Ɏg��
 
    ' outlook�I�u�W�F�N�g���쐬����Set�������
    Set outlookApp = CreateObject("Outlook.Application")
 
    ' nameSpace�I�u�W�F�N�g���擾����Set�������
    Set myNameSpace = outlookApp.GetNamespace("MAPI")
     
    ' �t�H���_�[�̎w��(GetDefaultFolder�̈������A�e�t�H���_�ɑΉ����Ă���)
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
     
    ' �t�H���_���J��(�E�B���h�E���\�������)
    myFolder.Display
     
    ' ��������Ɏc�����I�u�W�F�N�g�����
    Set outlookApp = Nothing
    Set myNameSpace = Nothing
    Set myFolder = Nothing
     
End Sub


Sub �Y�t�t�@�C���ۑ�()

    Dim objInbox As Object
    Dim objFolder As Object
    Dim strPath As String
    Dim i As Long
     
    ' �e��I�u�W�F�N�g�ϐ�
    Dim outlookApp As Outlook.Application    ' Outlook�I�u�W�F�N�g������
    Dim myNameSpace As Outlook.Namespace     ' nameSpace�I�u�W�F�N�g������
    Dim myFolder As Outlook.Folder           ' �t�H���_���w�肷��̂Ɏg��
 
    ' outlook�I�u�W�F�N�g���쐬����Set�������
    Set outlookApp = CreateObject("Outlook.Application")
 
    ' nameSpace�I�u�W�F�N�g���擾����Set�������
    Set myNameSpace = outlookApp.GetNamespace("MAPI")
     
    ' �t�H���_�[�̎w��(GetDefaultFolder�̈������A�e�t�H���_�ɑΉ����Ă���)
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
     
     
    Set objInbox = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    '�Y�t�t�@�C�������郁�[���̃t�H���_���w�肵�܂��B2�K�w�ȏ゠��ꍇ�́u.Folders.Item(���t�H���_����)�v��ǉ����Ă��������B
    Set objFolder = objInbox.Folders.Item("�l")
    
    '�Y�t�t�@�C���̕ۑ�����p�X�Ŏw�肵�܂��B
    strPath = "C:\Users\smuh1\Desktop\TestData\"
     
    For Each objItem In objFolder.Items
        For i = 1 To objItem.Attachments.Count
            '�Y�t�t�@�C���Ɋg���q������ꍇ�̂ݏ������܂��B
            If InStr(objItem.Attachments.Item(i), ".") <> 0 Then
                objItem.Attachments.Item(i).SaveAsFile strPath & objItem.Attachments.Item(i)
            End If
        Next i
    Next objItem
 
    Set objItem = Nothing
    Set objInbox = Nothing
    Set objFolder = Nothing
 
End Sub