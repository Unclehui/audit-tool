Attribute VB_Name = "ģ��1"
Sub һ��һ()
Attribute һ��һ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' һ��һ ��
'

'
   
    Application.DisplayAlerts = False
    
    Sheets.Add
    ActiveSheet.Name = "������һ��һ"
    Sheets("��").Select
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets("������һ��һ").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets.Add
    ActiveSheet.Name = "�˵���һ��һ"
    Sheets("��").Select
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets("�˵���һ��һ").Select
    Range("A1").Select
    ActiveSheet.Paste
    

    Dim i%
    Dim j%
       
    For i = 1 To Worksheets("�˵���һ��һ").Cells(65536, 1).End(xlUp).Value:
        For j = 1 To Worksheets("������һ��һ").Cells(65536, 1).End(xlUp).Value:
            If Round((Worksheets("�˵���һ��һ").Cells((i + 2), 2).Value) / 100) = Round((Worksheets("������һ��һ").Cells((j + 2), 2).Value) / 100) _
            And Round((Worksheets("�˵���һ��һ").Cells((i + 2), 7).Value - Worksheets("�˵���һ��һ").Cells((i + 2), 8).Value) * 100) _
            = Round((Worksheets("������һ��һ").Cells((j + 2), 4).Value - Worksheets("������һ��һ").Cells((j + 2), 3).Value) * 100) _
            And Worksheets("������һ��һ").Cells((j + 2), 12).Value = 0 Then
            Worksheets("�˵���һ��һ").Cells((i + 2), 12).Value = 1
            Worksheets("������һ��һ").Cells((j + 2), 12).Value = 1
            Worksheets("�˵���һ��һ").Cells((i + 2), 9).Value = Worksheets("������һ��һ").Cells((j + 2), 2).Value
            Worksheets("�˵���һ��һ").Cells((i + 2), 10).Value = Worksheets("������һ��һ").Cells((j + 2), 3).Value
            Worksheets("�˵���һ��һ").Cells((i + 2), 11).Value = Worksheets("������һ��һ").Cells((j + 2), 4).Value
            Worksheets("������һ��һ").Cells((j + 2), 5).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 2).Value
            Worksheets("������һ��һ").Cells((j + 2), 6).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 3).Value
            Worksheets("������һ��һ").Cells((j + 2), 7).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 4).Value
            Worksheets("������һ��һ").Cells((j + 2), 8).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 5).Value
            Worksheets("������һ��һ").Cells((j + 2), 9).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 6).Value
            Worksheets("������һ��һ").Cells((j + 2), 10).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 7).Value
            Worksheets("������һ��һ").Cells((j + 2), 11).Value = Worksheets("�˵���һ��һ").Cells((i + 2), 8).Value
            Exit For
            End If
        Next
    Next
    
    
    
    
    Application.DisplayAlerts = True

End Sub

Sub һ�Զ�()
'
' һ�Զ� ��
'

'
    Application.DisplayAlerts = False
    
    Sheets.Add
    ActiveSheet.Name = "������һ�Զ�"
    Sheets("������һ��һ").Select
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets("������һ�Զ�").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets.Add
    ActiveSheet.Name = "�˵���һ�Զ�"
    Sheets("�˵���һ��һ").Select
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets("�˵���һ�Զ�").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets.Add
    ActiveSheet.Name = "TEMP"
    Sheets.Add
    ActiveSheet.Name = "TEMP0"

    Sheets("������һ�Զ�").Select
    Range("A1").Select
    DNUM = Range("a65536").End(xlUp).Value
    Sheets("�˵���һ�Զ�").Select
    Range("A1").Select
    ZNUM = Range("a65536").End(xlUp).Value

    ' �˶Ե�

    Dim x%
    Dim calx As Integer
    calx = 2
    For x = 1 To Worksheets("�˵���һ�Զ�").Range("a65536").End(xlUp).Value:
        If Worksheets("�˵���һ�Զ�").Cells((x + 2), 12).Value = 0 Then
            Worksheets("TEMP0").Cells((calx), 1).Value = Worksheets("�˵���һ�Զ�").Cells((x + 2), 1).Value
            Worksheets("TEMP0").Cells((calx), 2).Value = Worksheets("�˵���һ�Զ�").Cells((x + 2), 7).Value - Worksheets("�˵���һ�Զ�").Cells((x + 2), 8).Value
            Worksheets("TEMP0").Cells((calx), 3).Value = Round((Worksheets("�˵���һ�Զ�").Cells((x + 2), 2).Value) / 100)
            Worksheets("TEMP0").Cells((calx), 4).Value = Abs(Worksheets("�˵���һ�Զ�").Cells((x + 2), 7).Value - Worksheets("�˵���һ�Զ�").Cells((x + 2), 8).Value)
            calx = calx + 1
        End If
    Next
    With Worksheets("TEMP0").Range("D1").CurrentRegion
        .Sort Key1:=.Range("D1"), order1:=xlAscending, Header:=xlYes
    End With
    Worksheets("TEMP0").Rows(1).Delete
    
    Dim i%
    For i = 1 To ZNUM:
        Sheets("TEMP").Select
        Cells.Select
        Selection.ClearContents
        Worksheets("TEMP").Range("A1").Value = Worksheets("TEMP0").Cells(i, 1).Value
        Worksheets("TEMP").Range("B1").Value = Worksheets("TEMP0").Cells(i, 2).Value
        Dim cal As Integer
        cal = 1
        Dim j%
        For j = 1 To DNUM:
            If Worksheets("TEMP0").Cells(i, 3).Value = Round((Worksheets("������һ�Զ�").Cells((j + 2), 2).Value) / 100) _
            And (Worksheets("TEMP0").Cells(i, 2).Value) / _
            (Worksheets("������һ�Զ�").Cells((j + 2), 4).Value - Worksheets("������һ�Զ�").Cells((j + 2), 3).Value) > 1 _
            And Worksheets("������һ�Զ�").Cells((j + 2), 12).Value = 0 Then
            Worksheets("TEMP").Cells(cal, 3).Value = Worksheets("������һ�Զ�").Cells((j + 2), 1).Value
            Worksheets("TEMP").Cells(cal, 4).Value = Worksheets("������һ�Զ�").Cells((j + 2), 4).Value - Worksheets("������һ�Զ�").Cells((j + 2), 3).Value
            cal = cal + 1
            End If
        Next
        If Worksheets("TEMP").Cells(1, 3).Value <> "" Then
            '�ж��Ƿ�Ϊ�ܺ�
            TEMPROWNUM = Worksheets("TEMP").Range("c65536").End(xlUp).Row
            Worksheets("TEMP").Range("F1").Value = "=SUMPRODUCT(D1:D" & "" & TEMPROWNUM & ",E1:E" & "" & TEMPROWNUM & ")"
            RANSTR = "$E$1:$E$" & "" & TEMPROWNUM
            CALNUM = Worksheets("TEMP").Range("B1").Value
            If Application.WorksheetFunction.Sum(Range(RANSTR)) = Worksheets("TEMP").Range("B1").Value Then
            Dim sumx%
            For sumx = 1 To TEMPROWNUM:
                Worksheets("TEMP").Cells(sumx, 5).Value = 1
            Next
            Worksheets("TEMP").Range("F1").Value = Application.WorksheetFunction.Sum(Range(RANSTR))
            Else:
                '�滮���
                SolverReset
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM, ByChange:=RANSTR _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverAdd CellRef:=RANSTR, Relation:=1, FormulaText:="1"
                SolverAdd CellRef:=RANSTR, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:=RANSTR, Relation:=4, FormulaText:="����"
                SolverOptions MaxSubproblems:=1073741824
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM, ByChange:=RANSTR _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM, ByChange:=RANSTR _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverSolve UserFinish:=True
            End If

            ' д��
            Dim MARKSTR As String
            MARKSTR = "������ţ�"
            Worksheets("TEMP").Range("F2").Value = Worksheets("TEMP").Range("F1").Value
            If Worksheets("TEMP").Range("F2").Value = Worksheets("TEMP").Range("B1").Value Then
                Dim w1%
                For w1 = 1 To TEMPROWNUM:
                    If Worksheets("TEMP").Cells(w1, 5).Value = 1 Then
                    DCODE = Worksheets("TEMP").Cells(w1, 3).Value
                    ZCODE = Worksheets("TEMP").Cells(1, 1).Value
                    MARKSTR = MARKSTR & "" & DCODE & ";"
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 5).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 2).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 6).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 3).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 7).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 4).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 8).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 5).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 9).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 6).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 10).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 7).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 11).Value = Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 8).Value
                    Worksheets("������һ�Զ�").Cells((DCODE + 2), 12).Value = 1
                    End If
                Next
                Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 13).Value = MARKSTR
                Worksheets("�˵���һ�Զ�").Cells((ZCODE + 2), 12).Value = 1
            End If
        End If
    Next
    
    Sheets("TEMP0").Select
    Cells.Select
    Selection.ClearContents

    '������

    Dim y%
    Dim caly As Integer
    caly = 2
    For y = 1 To DNUM:
        If Worksheets("������һ�Զ�").Cells((y + 2), 12).Value = 0 Then
            Worksheets("TEMP0").Cells((caly), 1).Value = Worksheets("������һ�Զ�").Cells((y + 2), 1).Value
            Worksheets("TEMP0").Cells((caly), 2).Value = Round(Worksheets("������һ�Զ�").Cells((y + 2), 4).Value - Worksheets("������һ�Զ�").Cells((y + 2), 3).Value)
            Worksheets("TEMP0").Cells((caly), 3).Value = Round((Worksheets("������һ�Զ�").Cells((y + 2), 2).Value) / 100)
            Worksheets("TEMP0").Cells((caly), 4).Value = Abs(Round(Worksheets("������һ�Զ�").Cells((y + 2), 4).Value - Worksheets("������һ�Զ�").Cells((y + 2), 3).Value))
            caly = caly + 1
        End If
    Next
    With Worksheets("TEMP0").Range("D1").CurrentRegion
        .Sort Key1:=.Range("D1"), order1:=xlAscending, Header:=xlYes
    End With
    Worksheets("TEMP0").Rows(1).Delete

    Dim m%
    For m = 1 To Worksheets("TEMP0").Range("a65536").End(xlUp).Row:
        Sheets("TEMP").Select
        Cells.Select
        Selection.ClearContents
        Worksheets("TEMP").Range("A1").Value = Worksheets("TEMP0").Cells(m, 1).Value
        Worksheets("TEMP").Range("B1").Value = Worksheets("TEMP0").Cells(m, 2).Value
        Dim cal2 As Integer
        cal2 = 1
        Dim n%
        For n = 1 To ZNUM:
            If Worksheets("TEMP0").Cells(m, 3).Value = Round((Worksheets("�˵���һ�Զ�").Cells((n + 2), 2).Value) / 100) _
            And (Worksheets("TEMP0").Cells(m, 2).Value) / _
            (Worksheets("�˵���һ�Զ�").Cells((n + 2), 7).Value - Worksheets("�˵���һ�Զ�").Cells((n + 2), 8).Value) > 1 _
            And Worksheets("�˵���һ�Զ�").Cells((n + 2), 12).Value = 0 Then
            Worksheets("TEMP").Cells(cal2, 3).Value = Worksheets("�˵���һ�Զ�").Cells((n + 2), 1).Value
            Worksheets("TEMP").Cells(cal2, 4).Value = Worksheets("�˵���һ�Զ�").Cells((n + 2), 7).Value - Worksheets("�˵���һ�Զ�").Cells((n + 2), 8).Value
            cal2 = cal2 + 1
            End If
        Next

        If Worksheets("TEMP").Cells(1, 3).Value <> "" Then
            TEMPROWNUM2 = Worksheets("TEMP").Range("c65536").End(xlUp).Row
            Worksheets("TEMP").Range("F1").Value = "=SUMPRODUCT(D1:D" & "" & TEMPROWNUM2 & ",E1:E" & "" & TEMPROWNUM2 & ")"
            RANSTR2 = "$E$1:$E$" & "" & TEMPROWNUM2
            CALNUM2 = Worksheets("TEMP").Range("B1").Value
            If Application.WorksheetFunction.Sum(Range(RANSTR2)) = Worksheets("TEMP").Range("B1").Value Then
            Dim sumy%
            For sumy = 1 To TEMPROWNUM2:
                Worksheets("TEMP").Cells(sumy, 5).Value = 1
            Next
            Worksheets("TEMP").Range("F1").Value = Application.WorksheetFunction.Sum(Range(RANSTR2))
            Else:
                TEMPROWNUM2 = Worksheets("TEMP").Range("c65536").End(xlUp).Row
                Worksheets("TEMP").Range("F1").Value = "=SUMPRODUCT(D1:D" & "" & TEMPROWNUM2 & ",E1:E" & "" & TEMPROWNUM2 & ")"
                RANSTR2 = "$E$1:$E$" & "" & TEMPROWNUM2
                CALNUM2 = Worksheets("TEMP").Range("B1").Value
                SolverReset
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM2, ByChange:=RANSTR2 _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverAdd CellRef:=RANSTR2, Relation:=1, FormulaText:="1"
                SolverAdd CellRef:=RANSTR2, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:=RANSTR2, Relation:=4, FormulaText:="����"
                SolverOptions MaxSubproblems:=1073741824
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM2, ByChange:=RANSTR2 _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverOk SetCell:="$F$1", MaxMinVal:=3, ValueOf:=CALNUM2, ByChange:=RANSTR2 _
                    , Engine:=1, EngineDesc:="Simplex LP"
                SolverSolve UserFinish:=True
            End If

            Dim MARKSTR2 As String
            MARKSTR2 = "������ţ�"
            Worksheets("TEMP").Range("F2").Value = Worksheets("TEMP").Range("F1").Value
            If Worksheets("TEMP").Range("F2").Value = Worksheets("TEMP").Range("B1").Value Then
                Dim w2%
                For w2 = 1 To TEMPROWNUM2:
                    If Worksheets("TEMP").Cells(w2, 5).Value = 1 Then
                    ZCODE2 = Worksheets("TEMP").Cells(w2, 3).Value
                    DCODE2 = Worksheets("TEMP").Cells(1, 1).Value
                    MARKSTR2 = MARKSTR2 & "" & ZCODE2 & ";"
                    Worksheets("�˵���һ�Զ�").Cells((ZCODE2 + 2), 9).Value = Worksheets("������һ�Զ�").Cells((DCODE2 + 2), 2).Value
                    Worksheets("�˵���һ�Զ�").Cells((ZCODE2 + 2), 10).Value = Worksheets("������һ�Զ�").Cells((DCODE2 + 2), 3).Value
                    Worksheets("�˵���һ�Զ�").Cells((ZCODE2 + 2), 11).Value = Worksheets("������һ�Զ�").Cells((DCODE2 + 2), 4).Value
                    Worksheets("�˵���һ�Զ�").Cells((ZCODE2 + 2), 12).Value = 1
                    End If
                Next
                Worksheets("������һ�Զ�").Cells((DCODE2 + 2), 13).Value = MARKSTR2
                Worksheets("������һ�Զ�").Cells((DCODE2 + 2), 12).Value = 1
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
    
    
End Sub
Sub ����ɾ��()
'
' ����ɾ�� ��
'

'
    Application.DisplayAlerts = False
    
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets:
        If sht.Name = "��" Or sht.Name = "��" Or sht.Name = "�����" Then
        
        Else: sht.Delete
        End If
    Next sht
    
    Application.DisplayAlerts = True
    

End Sub

