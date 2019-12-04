Attribute VB_Name = "GenerateEANFromValue"
Option Base 0

Sub DrawBar(BPosition As Double, BScale As Double, BHeight As Double, _
    alignX As Double, alignY As Double)
    'Call ActiveDocument.ActivePage.ActiveLayer.CreateRectangle(AlignX + BPosition, AlignY, _
    'AlignX + BScale + BPosition, AlignY + BHeight)
    Call ActiveDocument.ActivePage.ActiveLayer.CreateRectangle(alignX, alignY + BPosition, _
    alignX + BHeight, alignY + BScale + BPosition)
End Sub

Sub CreateSquareFromStamp()
    ' Запуск выбора и получение значения
    Call InputValueForm.Show
    If (Not InputValueForm.ValueSelected) Then Exit Sub

    v& = CLng(Val(InputValueForm.BarcodeValue.Text))

    Call Unload(InputValueForm)
    
    ' Создание таблицы значений EAN
    Dim L(10) As Integer
    L(0) = 13
    L(1) = 25
    L(2) = 19
    L(3) = 61
    L(4) = 35
    L(5) = 49
    L(6) = 47
    L(7) = 59
    L(8) = 55
    L(9) = 11
    
    Dim R(10) As Integer
    R(0) = 114
    R(1) = 102
    R(2) = 108
    R(3) = 66
    R(4) = 92
    R(5) = 78
    R(6) = 80
    R(7) = 68
    R(8) = 72
    R(9) = 116
    
' Dim G(10) As Integer
' 0   0001101 1110010 0100111
' 1   0011001 1100110 0110011
' 2   0010011 1101100 0011011
' 3   0111101 1000010 0100001
' 4   0100011 1011100 0011101
' 5   0110001 1001110 0111001
' 6   0101111 1010000 0000101
' 7   0111011 1000100 0010001
' 8   0110111 1001000 0001001
' 9   0001011 1110100 0010111
    
    ' Разделение входного значения и расчёт контрольной суммы
    Dim vd(8) As Long
    Dim cs As Integer
    cs = 0
    
    For i = 1 To 7
        vd(8 - i) = (v& Mod (10 ^ i)) \ (10 ^ (i - 1))
        If (i Mod 2 = 0) Then
           cs = cs + vd(8 - i)
        Else
           cs = cs + vd(8 - i) * 3
        End If
    Next
    vd(8) = (10 - (cs Mod 10)) Mod 10
    
    ' Получение поля отрисовки
    Dim alignX As Double: alignX = ActiveDocument.SelectionRange.LeftX
    Dim alignY As Double: alignY = ActiveDocument.SelectionRange.BottomY
    
    ' Отрисовка
    Dim leftPos As Double: leftPos = 0
    Dim bcScale As Double: bcScale = 0.015
    Dim bcHeight As Double: bcHeight = 25 * bcScale
    'Dim bcSideLinesHeight As Double: bcSideLinesHeight = 1#
    
    ' Калибровочные линии
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    leftPos = leftPos + bcScale
            
    ' Линии L/G
    For i = 1 To 4
        For j = 6 To 0 Step -1
            If (L(vd(i)) And (2 ^ j)) <> 0 Then
                Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
            End If
        
            leftPos = leftPos + bcScale
        Next j
    Next i
    
    ' Калибровочные линии
    leftPos = leftPos + bcScale
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    leftPos = leftPos + 2 * bcScale

    ' Линии R
    For i = 5 To 8
        For j = 6 To 0 Step -1
            If (R(vd(i)) And (2 ^ j)) <> 0 Then
                Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
            End If
        
            leftPos = leftPos + bcScale
        Next j
    Next i

    ' Калибровочные линии
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight, alignX, alignY)
    
    ' Завершено
End Sub
