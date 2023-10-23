Function 연관리금액(월관리금액, 개월수)

    yearly_price = 월관리금액 * 개월수
    yearly_price = Format(yearly_price, "#,#")
    term = "(" & 개월수 & "개월)"
    
    연관리금액 = yearly_price & term
    
End Function
