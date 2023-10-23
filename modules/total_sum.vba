Function 연납입금액(values As Range) As Double
    
    Dim c As Range
    Dim sum As Double
    
    For Each c In values
    
        ' 숫자인 경우엔 그냥 더하기
        If IsNumeric(c.value) = True Then
            sum = sum + c.value

        ' 숫자가 아니면 변환하기
        Else
            sum = sum + 금액추출(c.value)

        End If
    Next
    
    연납입금액 = sum
        
End Function
