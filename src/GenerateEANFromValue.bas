Attribute VB_Name = "GenerateEANFromValue"
Option Base 0

' Отрисовка одного столбца штрихкода
Sub DrawBar(BPosition As Double, BScale As Double, BHeight As Double, _
    AlignX As Double, AlignY As Double, Vertical As Boolean)
    If Vertical Then
        Call ActiveDocument.ActivePage.ActiveLayer.CreateRectangle(AlignX, AlignY + BPosition, _
            AlignX + BHeight, AlignY + BScale + BPosition)
    Else
        Call ActiveDocument.ActivePage.ActiveLayer.CreateRectangle(AlignX + BPosition, AlignY, _
            AlignX + BScale + BPosition, AlignY + BHeight)
    End If
End Sub

' Создать штрихкод EAN8
Sub CreateEAN8()
    Call CreateEAN(True)
End Sub

' Создать штрихкод EAN8
Sub CreateEAN13()
    Call CreateEAN(False)
End Sub

' Общая процедура создания штрихкода
Sub CreateEAN(EAN8 As Boolean)
    ' Переменные
    Dim endFactor As Integer
    If EAN8 Then endFactor = 8 Else endFactor = 13
    InputValueForm.BarcodeValue.MaxLength = endFactor - 1

    Dim leftPos As Double: leftPos = 0                          ' Текущая позиция отрисовки
    Dim bcScale As Double: bcScale = 0.015                      ' Масштаб изображения
    Dim bcHeight As Double: bcHeight = 30 * bcScale             ' Высота обычных полос
    Dim bcExtraHeight As Double: bcExtraHeight = 3 * bcScale    ' Высота калибровочных полос
    
    Dim vd(13) As Currency          ' Цифры штрихкода
    Dim cs As Integer: cs = 0       ' Контрольная сумма штрихкода
    
    Dim AlignX As Double: AlignX = ActiveDocument.SelectionRange.LeftX      ' Поле отрисовки
    Dim AlignY As Double: AlignY = ActiveDocument.SelectionRange.BottomY
    Dim Vert As Boolean: Vert = False                           ' Вертикальная ориентация ШК
    
    ' Запуск выбора и получение значения
    Call InputValueForm.Show
    If (Not InputValueForm.ValueSelected) Then Exit Sub

    Dim s As String: s = InputValueForm.BarcodeValue.Text

    Call Unload(InputValueForm)
    
    ' Создание таблицы кодов EAN
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
    
    Dim G(10) As Integer
    G(0) = 39
    G(1) = 51
    G(2) = 27
    G(3) = 33
    G(4) = 29
    G(5) = 57
    G(6) = 5
    G(7) = 17
    G(8) = 9
    G(9) = 23
    
    Dim C(10) As Integer
    C(0) = 0
    C(1) = 52 '11
    C(2) = 44 '13
    C(3) = 28 '14
    C(4) = 50 '19
    C(5) = 38 '25
    C(6) = 14 '28
    C(7) = 42 '21
    C(8) = 26 '22
    C(9) = 22 '26
   
    ' Разделение входного значения и расчёт контрольной суммы
    For i = 1 To endFactor - 1
        vd(i) = Val(Mid(s, i, 1))
        If (i Mod 2 = endFactor Mod 2) Then
           cs = cs + vd(i)
        Else
           cs = cs + vd(i) * 3
        End If
    Next
    vd(endFactor) = (10 - (cs Mod 10)) Mod 10
    
    ' Отрисовка
    
    ' Калибровочные линии
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    leftPos = leftPos + bcScale
            
    ' Линии L/G
    If EAN8 Then
        For i = 1 To endFactor \ 2
            ' Цифра
            Call ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(leftPos + AlignX, _
                AlignY - 2 * bcExtraHeight, Str(vd(i)))

            ' Полоса
            For j = 6 To 0 Step -1
                If (L(vd(i)) And (2 ^ j)) <> 0 Then
                    Call DrawBar(leftPos, bcScale, bcHeight, AlignX, AlignY, Vert)
                End If
            
                leftPos = leftPos + bcScale
            Next j
        Next i
        
    ' EAN13
    Else
        ' Первая цифра
        Call ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(AlignX - 7 * bcScale, _
            AlignY - 2 * bcExtraHeight, Str(vd(1)))
        
        For i = 1 To endFactor \ 2
            ' Цифра
            Call ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(leftPos + AlignX, _
                AlignY - 2 * bcExtraHeight, Str(vd(i + 1)))

            ' Полоса
            For j = 6 To 0 Step -1
                If ((C(vd(1)) And (2 ^ (i - 1))) <> 0) And ((G(vd(i + 1)) And (2 ^ j)) <> 0) Then
                    Call DrawBar(leftPos, bcScale, bcHeight, AlignX, AlignY, Vert)
                End If
                
                If ((C(vd(1)) And (2 ^ (i - 1))) = 0) And ((L(vd(i + 1)) And (2 ^ j)) <> 0) Then
                    Call DrawBar(leftPos, bcScale, bcHeight, AlignX, AlignY, Vert)
                End If
            
                leftPos = leftPos + bcScale
            Next j
        Next i
    End If
    
    ' Калибровочные линии
    leftPos = leftPos + bcScale
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    leftPos = leftPos + 2 * bcScale

    ' Линии R
    For i = endFactor \ 2 + 1 To (endFactor \ 2) * 2
        ' Цифра
        If EAN8 Then
            Call ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(leftPos + AlignX, _
                AlignY - 2 * bcExtraHeight, Str(vd(i)))
        Else
            Call ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(leftPos + AlignX, _
                AlignY - 2 * bcExtraHeight, Str(vd(i + 1)))
        End If
                
        ' Полоса
        For j = 6 To 0 Step -1
            If EAN8 And (R(vd(i)) And (2 ^ j)) <> 0 _
                Or Not EAN8 And (R(vd(i + 1)) And (2 ^ j)) <> 0 Then
                
                Call DrawBar(leftPos, bcScale, bcHeight, AlignX, AlignY, Vert)
            End If
        
            leftPos = leftPos + bcScale
        Next j
    Next i

    ' Калибровочные линии
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    leftPos = leftPos + 2 * bcScale
    Call DrawBar(leftPos, bcScale, bcHeight + bcExtraHeight, AlignX, AlignY - bcExtraHeight, Vert)
    
    ' Завершено
End Sub
