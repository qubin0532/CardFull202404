Attribute VB_Name = "ģ��1"

Sub ���û���ͳ�Ʊ��ʽ()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Set wb = Workbooks("000��ȫԺ��2024��04��У԰����ֵ���ܱ�(��ʽ).xlsx") ' �滻Ϊ��Ĺ���������
    Set ws = wb.Sheets("2024��04��")
    
    ' ѡ��A8:E45��Ԫ��
    Set rng = ws.Range("A8:e45")
    
    ' ����߿�����
    With rng
        .Borders.LineStyle = xlNone ' ����߿�
        .Interior.ColorIndex = xlNone ' ������
    End With
    
    ' ���ѡ��
    Set rng = Nothing
    
    
    ' �����и�
    ws.Rows("3:5").RowHeight = 5 'line 3-5 height=5
    ws.Rows("8:47").RowHeight = 27
    
    Columns("A:F").EntireColumn.AutoFit
    '�Զ������п�
    
    ' ѡ��A8:F33��Ԫ��
    Set rng = ws.Range("A8:e47")
    
    ' ����������ֺ�
    With rng.Font
        .Name = "����"
        .Size = 14
    End With
    
    ' ���ð����Ƶ�ɫ
    For Each cell In rng.Rows
        If cell.Row Mod 2 = 1 Then
            cell.Interior.Color = RGB(230, 230, 230) '��ɫ
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
    
    
    '���ж���ѡ�е�Ԫ��A��
    With rng.Columns("A")
        .HorizontalAlignment = xlCenter
    End With
    
    
    ' ���ж���C��D������
    With rng.Columns("C:D")
        .HorizontalAlignment = xlCenter
    End With
    
    ' ����E��Ϊ���ר�ø�ʽ
    rng.Columns("E").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    ' ������������ΪTimes New Roman
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            cell.Font.Name = "Times New Roman"
        End If
    Next cell
    
    ' ���ѡ��
    Set rng = Nothing
    
    
    ws.Rows("38").Hidden = True
    Set rng = ws.Range("7:7,40:47") 'ʹ��range���ԣ�ȡ��ѡ����ע�⣺ѡ�������ö��Ÿ�����
    
    rng.Rows.EntireRow.Hidden = True
    '���������ȡ�������������    rng = Range("7:7,40:48")
    
    'Rows("7:7").EntireRow.Hidden = True
    
    'Rows("32:48").EntireRow.Hidden = True
    
End Sub


