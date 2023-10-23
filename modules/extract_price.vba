Function 금액추출(price_with_format)

    ' 개월수를 떼고 앞부분 금액만 가져 옴
    priceWithComma = Split(price_with_format, "(")(0)

    ' 금액 중간 쉼표들을 제거
    priceWithoutComma = Replace(priceWithComma, ",", "")
    
    ' 문자열을 숫자로 바꿔서 반환
    금액추출 = Val(priceWithoutComma)
    
End Function