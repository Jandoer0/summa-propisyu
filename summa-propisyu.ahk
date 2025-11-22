; Скрипт для вставки суммы прописью
; Hotkey: Ctrl+Alt+W
^!w:: {
    try {
        ; Сохраняем оригинальный буфер обмена
        originalClipboard := A_Clipboard
        
        ; Ждем стабилизации буфера обмена
        Sleep(100)
        
        ; Проверяем, есть ли данные в буфере обмена
        if (A_Clipboard = "") {
            MsgBox "Буфер обмена пуст", "Ошибка", 0x10
            return
        }
        
        text := A_Clipboard
        if text = "" {
            MsgBox "Буфер обмена пуст", "Ошибка", 0x10
            return
        }
        
        ; Преобразуем сумму в прописной формат
        result := ConvertAmountToWords(text)
        
        ; Вставляем результат
        A_Clipboard := result
        Sleep(100) ; Даем время на запись в буфер
        
        Send("^v")
        ToolTip "Сумма вставлена прописью"
        SetTimer () => ToolTip(), -2000
        
        ; Восстанавливаем оригинальный буфер обмена через 2 секунды
        SetTimer () => A_Clipboard := originalClipboard, -2000
        
    } catch Error as e {
        MsgBox "Неверный формат числа.`n`nПоддерживаемые форматы:`n• 1234567,89`n• 1 234 567,89`n• 1234567`n• 123,45", "Ошибка преобразования", 0x10
        ; Восстанавливаем оригинальный буфер обмена при ошибке
        try A_Clipboard := originalClipboard
    }
}

ConvertAmountToWords(amount) {
    ; Очищаем входную строку от пробелов и лишних символов
    amount := Trim(amount)
    amount := StrReplace(amount, " ", "")
    amount := StrReplace(amount, "`t", "")
    
    ; Проверяем формат числа (более гибкая проверка)
    if !RegExMatch(amount, "^(?:\d{1,3}(?: \d{3})*|\d+)(?:[.,]\d{1,2})?$") && !RegExMatch(amount, "^\d+([.,]\d{1,2})?$") {
        throw Error("Неправильный формат числа")
    }
    
    ; Заменяем запятую на точку для удобства обработки
    amount := StrReplace(amount, ",", ".")
    
    ; Разделяем на рубли и копейки
    parts := StrSplit(amount, ".")
    rubles := parts[1]
    kopecks := parts.Length > 1 ? parts[2] : "00"
    
    ; Дополняем копейки до 2 знаков
    if StrLen(kopecks) = 1 {
        kopecks .= "0"
    } else if StrLen(kopecks) > 2 {
        kopecks := SubStr(kopecks, 1, 2)
    }
    
    ; Преобразуем рубли в текст
    rublesText := ConvertNumberToWords(rubles)
    
    ; Определяем правильную форму слова "рубль"
    rublesForm := GetCurrencyForm(rubles, "рубль", "рубля", "рублей")
    
    ; Формируем окончательный результат
    result := rublesText " " rublesForm " " kopecks " коп."
    
    return StrReplace(result, "  ", " ") ; Убираем двойные пробелы
}

ConvertNumberToWords(number) {
    if number = "0" {
        return "ноль"
    }
    
    ; Массивы для преобразования чисел в слова
    units := ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    teens := ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семьнадцать", "восемнадцать", "девятнадцать"]
    tens := ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
    hundreds := ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"]
    
    ; Классы чисел
    classes := [
        ["", "", ""],
        ["тысяча", "тысячи", "тысяч"],
        ["миллион", "миллиона", "миллионов"],
        ["миллиард", "миллиарда", "миллиардов"]
    ]
    
    ; Разбиваем число на классы (по 3 цифры)
    numStr := Format("{:01}", number) ; Убираем ведущие нули
    
    ; Дополняем нулями слева до количества, кратного 3
    length := StrLen(numStr)
    padding := Mod(3 - Mod(length, 3), 3)
    if padding > 0 {
        numStr := Format("{:0" (length + padding) "}", number)
    }
    
    classCount := StrLen(numStr) // 3
    result := ""
    
    Loop classCount {
        classIndex := classCount - A_Index + 1
        startPos := (A_Index - 1) * 3 + 1
        classStr := SubStr(numStr, startPos, 3)
        classNum := Integer(classStr)
        
        if classNum = 0 {
            Continue
        }
        
        ; Преобразуем трехзначное число
        classText := ""
        
        ; Сотни
        hundredsDigit := SubStr(classStr, 1, 1)
        if hundredsDigit != "0" {
            classText .= hundreds[Integer(hundredsDigit) + 1] " "
        }
        
        ; Десятки и единицы
        tensDigit := SubStr(classStr, 2, 1)
        unitsDigit := SubStr(classStr, 3, 1)
        
        if tensDigit = "1" {
            classText .= teens[Integer(unitsDigit) + 1] " "
        } else {
            if tensDigit != "0" {
                classText .= tens[Integer(tensDigit) + 1] " "
            }
            if unitsDigit != "0" {
                ; Особые формы для тысяч
                if classIndex = 2 {
                    if unitsDigit = "1" {
                        classText .= "одна "
                    } else if unitsDigit = "2" {
                        classText .= "две "
                    } else {
                        classText .= units[Integer(unitsDigit) + 1] " "
                    }
                } else {
                    classText .= units[Integer(unitsDigit) + 1] " "
                }
            }
        }
        
        ; Добавляем название класса
        if classIndex > 1 && classNum > 0 {
            classForm := GetCurrencyForm(classNum, classes[classIndex][1], classes[classIndex][2], classes[classIndex][3])
            classText .= classForm " "
        }
        
        result .= classText
    }
    
    return Trim(result)
}

GetCurrencyForm(number, form1, form2, form3) {
    ; Определяем правильную форму слова в зависимости от числа
    n := Mod(number, 100)
    if n >= 11 && n <= 19 {
        return form3
    }
    
    n := Mod(number, 10)
    if n = 1 {
        return form1
    } else if n >= 2 && n <= 4 {
        return form2
    } else {
        return form3
    }
}