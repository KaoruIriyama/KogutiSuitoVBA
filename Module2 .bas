Attribute VB_Name = "Module2"
Option Explicit
Sub ����Ȗڂ��Ƃ̕\�̋��z���W�v�����v�\�ɓ]�L����()
    Dim X As Long
    Dim Y As Long
    Dim Z As Range
    
    For X = 5 To Sheets.Count '����Ȗڂ��Ƃ̃V�[�g�̐��������̏���������
        Sheets(X).Activate
        Sheets(X).Cells(2, 11) = WorksheetFunction.Sum(Sheets(X).Range(Range("G3"), Range("G3").End(xlDown)))          'S�Z���̍��v�𓯂��V�[�g�́u���v�v���ɕ\��
    Next X
    
    For Y = 5 To Sheets.Count '����Ȗڂ��Ƃ̃V�[�g�̐��������̏���������
        Sheets(Y).Activate
        Sheets(4).Cells(Y, 2) = Sheets(Y).Name         '�������̊���ȖڃV�[�g�̂����A���[�̂��̖̂��O���u���v�v�V�[�g�̖��O���ɃR�s�[
        Sheets(4).Cells(Y, 3) = Sheets(Y).Range("K2")  '�������̊���ȖڃV�[�g�̂����A���[�̂��̂́u���v�v�Z�����u���v�v�V�[�g�̖��O���ɃR�s�[
    Next Y
    
    Sheets(4).Activate
        Set Z = Sheets(4).Range(Range("C5"), Range("C5").End(xlDown)) '�ϐ�Z�ɓ]�L���ꂽ�Ȗڂ��Ƃ̍��v�Z���̗����
        Sheets(4).Cells(18, 6) = Application.WorksheetFunction.Sum(Z)  'Z�Z���̑��v�𓯂��V�[�g�́u���v�v���ɕ\��
    
    MsgBox "���v�̓��͂��������܂����B"
End Sub

