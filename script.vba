Sub Прямоугольникскругленныеуглы1_Щелчок()

'отключаем     всплывающие окна
Application.DisplayAlerts = False

'Счетчики
Dim BankIter As Integer, FactorIter As Integer

'Текущая книга
Set ThisWB = ThisWorkbook

'новая книга
Set NewWB = Workbooks.Add

'Листы текущей книги
Set Banks = ThisWB.Sheets

Set Factors = ThisWB.Sheets(1).Range("A3:A18")
Set Blocks = ThisWB.Sheets(1).Range("B1:Y1")

'Создание листов по факторам
For Each Factor In Factors
    NewWB.Sheets.Add(After:=NewWB.Sheets(NewWB.Sheets.Count)).Name = Left(Replace(Factor.Value, ":", "."), 30)
Next
'Удаляем лист по умолчанию
NewWB.Sheets(1).Delete

'Проход по каждому новому листу
For Each Sheet In NewWB.Sheets
'Вставка Отделений
    BankIter = 1
    For Each Bank In Banks
        Sheet.Cells((2 + BankIter), 1) = Bank.Name
        ThisWB.Sheets(1).Range("A3").Copy
        Sheet.Cells((2 + BankIter), 1).PasteSpecial xlPasteFormats
        BankIter = BankIter + 1
    Next
    
'Вставка шапки блоков
    Blocks.Copy
    Sheet.Range("B1").PasteSpecial xlPasteValues
    Sheet.Range("B1").PasteSpecial xlPasteFormats
Next

FactorIter = 3
BankIter = 3

'Транспонирование
For Each NewSheet In NewWB.Sheets
    For Each Bank In Banks
        Bank.Range("B" & BankIter & ":Y" & BankIter).Copy
        NewSheet.Range("B" & FactorIter).PasteSpecial xlPasteValues
        NewSheet.Range("B" & FactorIter).PasteSpecial xlPasteFormats
    FactorIter = FactorIter + 1
    Next
FactorIter = 3
BankIter = BankIter + 1
Next

'включаем     всплывающие окна
Application.DisplayAlerts = True
'Встаем на первый лист
NewWB.Sheets(1).Activate
End Sub
