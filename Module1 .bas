Attribute VB_Name = "Module1"
Option Explicit
Sub ����Ȗڂ��ƂɕʃV�[�g���o()

Dim Kanjoukamoku_List As Worksheet
Dim Data As Worksheet
Dim Genshi As Worksheet
Dim i As Long
Dim J As Long
Dim RowCnt  As Long
Dim LastRow  As Long
Dim List_Cnt As Long
Dim ShtName As String

    '�e�V�[�g��ϐ����
    Set Kanjoukamoku_List = Sheets("����Ȗڃ��X�g")
    Set Data = Sheets("���������o�[��")
    Set Genshi = Sheets("����")

    '�u����Ȗځv �ŏI�s
    List_Cnt = Kanjoukamoku_List.Cells(Rows.Count, 1).End(xlUp).Row

    '�u�f�[�^�v �ŏI�s
    LastRow = Data.Cells(Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

        '�u����Ȗڃ��X�g�v�����ƂɐV�K�V�[�g�쐬
        For i = 2 To List_Cnt
            Genshi.Copy after:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = Kanjoukamoku_List.Cells(i, 1)
        Next i

        '���Ӑ�̐������[�v
        For i = 2 To List_Cnt

            Data.Select

            '�\��t���J�n�s
            RowCnt = 3

            '�V�[�g���擾
            ShtName = Kanjoukamoku_List.Cells(i, 1)

            '�u�f�[�^�v 2�`�ŏI�s�܂Ń��[�v
            For J = 6 To LastRow

                '�u�f�[�^�v �� �u����Ȗڃ��X�g�v �Ɠ������̂���������
                If Kanjoukamoku_List.Cells(i, 1) = Data.Cells(J, 3) Then

                    '�u�f�[�^�v  A�`D��̒l���A�Y���V�[�g�ɓ\��t��
                    Data.Range(Cells(J, 1), Cells(J, 7)).Copy Sheets(ShtName).Cells(RowCnt, 1)

                    '�\��t���J�n�s���X�V
                    RowCnt = RowCnt + 1
               
                End If
            Next J
        Next i

    Application.ScreenUpdating = True
    
    MsgBox "����Ȗڂ��Ƃ̕\���쐬����܂����B�e�V�[�g�̓��e���`�F�b�N���A���M�C�������˂������܂��B"

End Sub


