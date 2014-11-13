Attribute VB_Name = "oRulesTool"
''Объявление переменной message, содержащей информацию о непрочтенных письмах
Public message As String

''Процедура выполнающая правила в outlook
Public Sub RulesRun()

''Обьявление переменных
Dim MyRules As Rules
Dim check As Byte

''Установка переменных
Set MyRules = Application.Session.DefaultStore.GetRules()
Rule = 1

''Цикл перебора правил
For Each Rule In MyRules
    
    ''Здесь проверяется, чтобы номер выполняемого правила не оказасся болше их общего количества
    If Rule <= MyRules.Count Then
        
        ''Выполенние правила
        MyRules.Item(Rule).Execute
        
        ''Счетчик правил увеличивается на 1
        Rule = Rule + 1
DoEvents
     End If

Next

''Вывод информационного сообщения о текущей ситуации с непрочитанными письмами
UFcontent

End Sub

''Процедура формирующая содержимое формы UserForm1
Public Sub UFcontent()

''Обьявление переменных
Dim oNamespace As Outlook.NameSpace
Dim oChildFolder As Outlook.MAPIFolder

''Установка переменных
Set oNamespace = Application.GetNamespace("MAPI")
Set oMailFolder = oNamespace.Folders("nikolai.karpov@heineken.com")
message = ""

''Проверка непрочитанных писем
CheckNM oMailFolder

''Заполнение Label1 в UserForm1 информацией о непрочитанных письмах
UserForm1.Label1.Caption = message

''Загрузка и отрисовка формы
UserForm1.Show

''Обратный таймер до закрытия формы, в скобках начальное значение
oTimer (10)

''Выгрузка формы
Unload UserForm1

End Sub

''Процедура поиска непрочитанных писем
Public Sub CheckNM(ByVal ofolder As MAPIFolder)

''Обьявление переменных
Dim oChildFolder As Outlook.MAPIFolder

''Цикл перебирающий все папки в корневой папке
For Each oChildFolder In ofolder.Folders

    ''Для каждой из папок проводится проверка наличая непрочитанных писем
    If oChildFolder.UnReadItemCount <> 0 Then
        ''Если непрочитанные письма найдены - добавление информации о них в переменную message
        message = message + oChildFolder & " - " & oChildFolder.UnReadItemCount & vbCrLf
    End If
    DoEvents
    ''запуска процедуры CheckNM для каждой из подпапок
    CheckNM oChildFolder
    
Next

End Sub

''Процедура выводящая таймер обратного отсчета начиная с Seconds, на данный момент в Label2
Public Sub oTimer(Seconds)

''Цикл обновления таймера
Do While Seconds > 0
    ''Присваивается значение поля Caption у Label2
    UserForm1.Label2.Caption = Seconds
    
    ''Перерисовывается форма
    UserForm1.Repaint
    
    ''VB замориживаетя на 1 сек
    SleepVB (1)
    
    ''Значение Second уменьшается на единицу
    Seconds = Seconds - 1

Loop

End Sub

''процедура приостанавливающая выполнение макроса на количество секунд указанное в Seconds
Sub SleepVB(Seconds)

''Обьявление переменных
Dim Start

''Установка Start равным текущиму времени в секундах
Start = timer

''Выполнять цекл до тех пор, пока текущее время не станет равным изначальному + Second
Do While timer < Start + Seconds

''Предача управления ОС
DoEvents

Loop

End Sub
