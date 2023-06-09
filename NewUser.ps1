<#----------------------------------
Имя сценария: NewUser.ps1
Дата создания: 04.10.2021
Дата последнего изменения: 12.05.2022
Версия: 1.0.2
*****************************
Описание: 1) Скрипт обращается к файлу указанному в переменной $File. 
        2) Читает данные из файла Excel с помощью PowerShell. 
        3) Проверяет существование пользователя
        4) Генерирует логин и переводит его в транслит
        5) Записывает логин и пароль из переменной $Pass в файл построчно
        6) Сохраняет Excel файл
----------------------------------#>

$LogDate = get-date -f yyyyMMdd
$LogFile = "C:\TEMP\New_Account_$LogDate.log"
$File = $null

# В которой создавать учётки
$OU = "OU=New Users,OU=Domain Users,DC=contoso,DC=LOCAL"

Import-Module ActiveDirectory

# Функция транслитерации кириллических символов в латиницу
function Translit
{
param([string]$inString)

    #Создаем хэш-таблицу соответствия русских и латинских символов на основе транслитерации, которую использует Википедия
    # https://en.wikipedia.org/wiki/Wikipedia:Romanization_of_Russian
    $Translit = @{

    [char]'а' = "a"
    [char]'А' = "A"
    [char]'б' = "b"
    [char]'Б' = "B"
    [char]'в' = "v"
    [char]'В' = "V"
    [char]'г' = "g"
    [char]'Г' = "G"
    [char]'д' = "d"
    [char]'Д' = "D"
    [char]'е' = "e"
    [char]'Е' = "Ye"
    [char]'ё' = "yo"
    [char]'Ё' = "Yo"
    [char]'ж' = "zh"
    [char]'Ж' = "Zh"
    [char]'з' = "z"
    [char]'З' = "Z"
    [char]'и' = "i"
    [char]'И' = "I"
    [char]'й' = "y"
    [char]'Й' = "Y"
    [char]'к' = "k"
    [char]'К' = "K"
    [char]'л' = "l"
    [char]'Л' = "L"
    [char]'м' = "m"
    [char]'М' = "M"
    [char]'н' = "n"
    [char]'Н' = "N"
    [char]'о' = "o"
    [char]'О' = "O"
    [char]'п' = "p"
    [char]'П' = "P"
    [char]'р' = "r"
    [char]'Р' = "R"
    [char]'с' = "s"
    [char]'С' = "S"
    [char]'т' = "t"
    [char]'Т' = "T"
    [char]'у' = "u"
    [char]'У' = "U"
    [char]'ф' = "f"
    [char]'Ф' = "F"
    [char]'х' = "kh"
    [char]'Х' = "Kh"
    [char]'ц' = "ts"
    [char]'Ц' = "Ts"
    [char]'ч' = "ch"
    [char]'Ч' = "Ch"
    [char]'ш' = "sh"
    [char]'Ш' = "Sh"
    [char]'щ' = "shch"
    [char]'Щ' = "Shch"
    [char]'ъ' = "" # "``"
    [char]'Ъ' = "" # "``"
    [char]'ы' = "y" # "y`"
    [char]'Ы' = "Y" # "Y`"
    [char]'ь' = "" # "`"
    [char]'Ь' = "" # "`"
    [char]'э' = "e" # "e`"
    [char]'Э' = "E" # "E`"
    [char]'ю' = "yu"
    [char]'Ю' = "Yu"
    [char]'я' = "ya"
    [char]'Я' = "Ya"
    [char]' ' = " " #пробел

    }

# обработка в цикле кириллических символов с заменой на латиницу
$outChars=""
foreach ($c in $inChars = $inString.ToCharArray()){
    if ($Translit[$c] -cne $Null ){
    $outChars += $Translit[$c]}
    else{
$outChars += $c}
}

Write-Output $outChars
}
# конец функции

# Функция генерации пароля
function GenPass {
    #param([string]$PassF)
        $PassA = -join ( % { [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ' | Get-Random })
        $PassB = -join (2..$PassLength | % { [char[]]'0123456789' | Get-Random })
        $PassC = -join ( % { [char[]]'-_.+=()' | Get-Random })
        $PassD = -join ( % { [char[]]'abcdefghijklmnopqrstuvwxyz' | Get-Random })
        $PassE = -join (2..$PassLength | % { [char[]]'0123456789' | Get-Random })
        $PassF = $PassA+$PassB+$PassC+$PassD+$PassE
        Write-Output $PassF
    }
# конец функции генерации пароля


#Рисуем GUI поле
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Web

    $window_form = New-Object System.Windows.Forms.Form
    $window_form.Text ='Создание нового пользователя SZM'
    $window_form.Width = 350
    $window_form.Height = 440
    $window_form.AutoSize = $true
    
#Рисуем поле для кнопок "Создать много пользователей"
    #Рисуем GUI кнопки "Выбрать файл"
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Point(5,25)
    $AddButton.Size = New-Object System.Drawing.Size(120,70)
    $AddButton.Text = 'Выбрать файл'
    #Действие при нажатии кнопки "Выбрать файл"
    $AddButton.Add_Click(
        {
        $f = new-object Windows.Forms.OpenFileDialog
#        $f.InitialDirectory = pwd
        $f.Filter = "Книга Excel (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $f.Multiselect = $true
        [void]$f.ShowDialog()
        if ($f.Multiselect) { $f.FileNames } else { $f.FileName }

        # Путь до файла
        $Global:File = ($f.FileName)
        Write-Host $File
        }
    )

#Рисуем GUI поле для ввода "Группа AD"
    $FormLabelb1 = New-Object System.Windows.Forms.Label
    $FormLabelb1.Text = "И поместить их в группу AD"
    $FormLabelb1.Location = New-Object System.Drawing.Point(135,35)
    $FormLabelb1.AutoSize = $true
    $TextBoxb1 = New-Object system.Windows.Forms.TextBox
    $TextBoxb1.multiline = $false
    $TextBoxb1.Size = New-Object System.Drawing.Size(150,20)
    $TextBoxb1.location = New-Object System.Drawing.Point(135,60)
    $TextBoxb1.Text = "RAS Users"

$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Location = New-Object System.Drawing.Size(5,10) 
$groupBox1.size = New-Object System.Drawing.Size(300,120)
$groupBox1.text = "Создать много пользователей из файла EXCEL:"  
$groupBox1.Controls.AddRange(@($FormLabelb1,$TextBoxb1, $AddButton))
    
    
#Рисуем поле для кнопок "Создать одного пользователя"
    #2
    $FormLabel2 = New-Object System.Windows.Forms.Label
    $FormLabel2.Text = "Фамилия"
    $FormLabel2.Location = New-Object System.Drawing.Point(10,20)
    $FormLabel2.AutoSize = $true
    $TextBox2 = New-Object system.Windows.Forms.TextBox
    $TextBox2.multiline = $false
    $TextBox2.Size = New-Object System.Drawing.Size(150,20)
    $TextBox2.location = New-Object System.Drawing.Point(125,20)
    #3
    $FormLabel3 = New-Object System.Windows.Forms.Label
    $FormLabel3.Text = "Имя"
    $FormLabel3.Location = New-Object System.Drawing.Point(10,45)
    $FormLabel3.AutoSize = $true
    $TextBox3 = New-Object system.Windows.Forms.TextBox
    $TextBox3.multiline = $false
    $TextBox3.Size = New-Object System.Drawing.Size(150,20)
    $TextBox3.location = New-Object System.Drawing.Point(125,45)
    #4
    $FormLabel4 = New-Object System.Windows.Forms.Label
    $FormLabel4.Text = "Отчество"
    $FormLabel4.Location = New-Object System.Drawing.Point(10,70)
    $FormLabel4.AutoSize = $true
    $TextBox4 = New-Object system.Windows.Forms.TextBox
    $TextBox4.multiline = $false
    $TextBox4.Size = New-Object System.Drawing.Size(150,20)
    $TextBox4.location = New-Object System.Drawing.Point(125,70)
    #6
    $FormLabel5 = New-Object System.Windows.Forms.Label
    $FormLabel5.Text = "Должность"
    $FormLabel5.Location = New-Object System.Drawing.Point(10,115)
    $FormLabel5.AutoSize = $true
    $TextBox5 = New-Object system.Windows.Forms.TextBox
    $TextBox5.multiline = $false
    $TextBox5.Size = New-Object System.Drawing.Size(150,20)
    $TextBox5.location = New-Object System.Drawing.Point(125,115)
    #6
    $FormLabel6 = New-Object System.Windows.Forms.Label
    $FormLabel6.Text = "Учреждение"
    $FormLabel6.Location = New-Object System.Drawing.Point(10,140)
    $FormLabel6.AutoSize = $true
    $TextBox6 = New-Object system.Windows.Forms.TextBox
    $TextBox6.multiline = $false
    $TextBox6.Size = New-Object System.Drawing.Size(150,20)
    $TextBox6.location = New-Object System.Drawing.Point(125,140)
    #7
    $FormLabel7 = New-Object System.Windows.Forms.Label
    $FormLabel7.Text = "Номер телефона"
    $FormLabel7.Location = New-Object System.Drawing.Point(10,165)
    $FormLabel7.AutoSize = $true
    $TextBox7 = New-Object system.Windows.Forms.TextBox
    $TextBox7.multiline = $false
    $TextBox7.Size = New-Object System.Drawing.Size(150,20)
    $TextBox7.location = New-Object System.Drawing.Point(125,165)
    #8
    $FormLabel8 = New-Object System.Windows.Forms.Label
    $FormLabel8.Text = "e-mail"
    $FormLabel8.Location = New-Object System.Drawing.Point(10,190)
    $FormLabel8.AutoSize = $true
    $TextBox8 = New-Object system.Windows.Forms.TextBox
    $TextBox8.multiline = $false
    $TextBox8.Size = New-Object System.Drawing.Size(150,20)
    $TextBox8.location = New-Object System.Drawing.Point(125,190)
    #9
    $FormLabel9 = New-Object System.Windows.Forms.Label
    $FormLabel9.Text = "Группа AD"
    $FormLabel9.Location = New-Object System.Drawing.Point(10,215)
    $FormLabel9.AutoSize = $true
    $TextBox9 = New-Object system.Windows.Forms.TextBox
    $TextBox9.multiline = $false
    $TextBox9.Size = New-Object System.Drawing.Size(150,20)
    $TextBox9.location = New-Object System.Drawing.Point(125,215)
    $TextBox9.Text = "New Users"
    
#Рисуем GUI "Создать одного пользователя"
$groupBox2 = New-Object System.Windows.Forms.GroupBox
$groupBox2.Location = New-Object System.Drawing.Size(5,180) 
$groupBox2.size = New-Object System.Drawing.Size(300,240)
$groupBox2.text = "Создать одного пользователя:"
$groupBox2.Controls.AddRange(@($FormLabelb2,$TextBoxb2,$FormLabel1,$TextBox1,$FormLabel2,$TextBox2,$FormLabel3,$TextBox3,$FormLabel4,$TextBox4,$FormLabel5,$TextBox5,$FormLabel6,$TextBox6,$FormLabel7,$TextBox7,$FormLabel8,$TextBox8,$FormLabel9,$TextBox9))

    $Or = New-Object System.Windows.Forms.Label
    $Or.Text = "--- ИЛИ ---"
    $Or.Location = New-Object System.Drawing.Point(130,150)
    $Or.AutoSize = $true


#Рисуем GUI кнопки "Выполнить скрипт"
$GoButton = New-Object System.Windows.Forms.Button
$GoButton.Location = New-Object System.Drawing.Point(310,155)
$GoButton.Size = New-Object System.Drawing.Size(110,100)
$GoButton.Text = 'Выполнить скрипт'
#Действие при нажатии кнопки
$GoButton.Add_Click(
        {
        $Surname = ($TextBox2.Text)
        $SurnameBox = ($TextBox2.Text)
        $GivenName = ($TextBox3.Text)
        $Patronymic = ($TextBox4.Text)
        $Title = ($TextBox5.Text)
        $Company = ($TextBox6.Text)
        $mobile = ($TextBox7.Text)
        $email = ($TextBox8.Text)
        
        #Проверка какое поле "Группа" импользовать. Если заполнено поле Фамилия или нет
        if ($SurnameBox) {
            $Group = ($TextBox9.Text)
            $FullName = "$Surname" + " " + "$GivenName" + " " + "$Patronymic"
            #Write-Host ″$Group″
            Write-Host "$FullName"

            # Генерируем логин
            $SamAccountNameRU = $Surname + "_" + $GivenName[0] + $Patronymic[0]
            $TransSamAccountNameEN = Translit($SamAccountNameRU)
            
            # Генерируем пароль
            $Pass = GenPass

            #Записываем данные в файл
            "$FullName" >> $LogFile
            "$TransSamAccountNameEN" >> $LogFile
            "$Pass" >> $LogFile
            " " >> $LogFile
            
            New-ADUser -Name "$FullName" -DisplayName "$FullName" -GivenName $GivenName -Surname $Surname -UserPrincipalName "$TransSamAccountNameEN@contoso.local" -SamAccountName $TransSamAccountNameEN -Company "$Company" -EmailAddress "$Email" -MobilePhone "$Mobile" -Title "$Title" -Path $OU -AccountPassword (ConvertTo-SecureString $Pass -AsPlainText -Force) -Enabled $true -ChangePasswordAtLogon $true
            Start-Sleep -Seconds 5
            Add-ADGroupMember -Identity $Group -Members $TransSamAccountNameEN

             # Выводим на экран информационное сообщение
            Write-Host "$TransSamAccountNameEN"
            Write-Host "$Pass" -f yellow 
            Write-Host " "
            }
        else {
            $Group = ($TextBoxb1.Text)
            Write-Host ″$Group″
            }


        if ($File) {
        # Для запуска приложения создаем COM-объект и помещаем его в переменную
        $Excel = New-Object -ComObject Excel.Application
        # Делаем его видимым
        $Excel.Visible = $true
        # Открываем книгу
        $WorkBook = $Excel.Workbooks.Open($File)
        #Выберем лист с именем "Лист1"
        $WorkSheet = $WorkBook.Sheets.Item("Лист1")
        # Добавляем новые столбцы
        $WorkSheet.Cells.Item(8,40) = 'LOGIN'
        $WorkSheet.Cells.Item(8,40).Font.Bold = $true
        $WorkSheet.Cells.Item(8,41) = 'PASSWORD'
        $WorkSheet.Cells.Item(8,41).Font.Bold = $true
        # Выделяем жирным шапку таблицы, для единства стиля
        $WorkSheet.Rows.Item(8).Font.Bold = $true
        #Кол-во занятых строк в файле
        $MaxRows = ($WorkSheet.UsedRange.Rows).count;
        #Кол-во занятых столбцов в файле
        $MaxColumns = ($WorkSheet.UsedRange.Columns).count;
        # Переменные для записи в файл
        $ColumnL = 40
        $ColumnP = 41
        # Парсим данные из файла
        $users=@();

        for ($row = 12; $row -le $MaxRows; $row++) {
        $user = New-Object -TypeName PSObject;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,1).Text -Value $WorkSheet.UsedRange.Cells($row,1).Text -MemberType NoteProperty;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,8).Text -Value $WorkSheet.UsedRange.Cells($row,8).Text -MemberType NoteProperty;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,7).Text -Value $WorkSheet.UsedRange.Cells($row,7).Text -MemberType NoteProperty;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,6).Text -Value $WorkSheet.UsedRange.Cells($row,6).Text -MemberType NoteProperty;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,11).Text -Value $WorkSheet.UsedRange.Cells($row,11).Text -MemberType NoteProperty;
        $user | Add-Member -Name $WorkSheet.UsedRange.Cells(8,12).Text -Value $WorkSheet.UsedRange.Cells($row,12).Text -MemberType NoteProperty;

        #Парсим
        $Surname = ($user."Фамилия ")
        $GivenName = ($user."Имя ")
        $Patronymic = ($user.Отчество)
        $Title = ($user."Должность  (бухгалтер по зп, экономист по труду и т.п.)")
        $Company = ($user.Учреждение)
        $mobile = ($user."Телефон рабочий")
        $FullName = "$Surname" + " " + "$GivenName" + " " + "$Patronymic"


        #Стопкран
        If (!$Surname){break}

        Write-Host "$FullName"
        
        # Генерируем логин
        $SamAccountNameRU = $Surname + "_" + $GivenName[0] + $Patronymic[0]
        $TransSamAccountNameEN = Translit($SamAccountNameRU)

        # Генерируем пароль        
        $Pass = GenPass

        #Заполняем столбец LOGIN И PASSWORD
        $WorkSheet.Cells.Item($row, $ColumnL) = $TransSamAccountNameEN
        $WorkSheet.Cells.Item($row, $ColumnP) = $Pass
        
        # Выводим на экран информационное сообщение
        Write-Host "$TransSamAccountNameEN"
        Write-Host "$Pass" -f yellow
        Write-Host " "

        # Создаем учётку и добавляем в группу
        New-ADUser -Name "$FullName" -DisplayName "$FullName" -GivenName $GivenName -Surname $Surname -UserPrincipalName "$TransSamAccountNameEN@contoso.local" -SamAccountName $TransSamAccountNameEN -Company "$Company" -EmailAddress "$Email" -MobilePhone "$Mobile" -Title "$Title" -Path $OU -AccountPassword (ConvertTo-SecureString $Pass -AsPlainText -Force) -Enabled $true -ChangePasswordAtLogon $true
        Add-ADGroupMember -Identity $Group -Members $TransSamAccountNameEN
        }
    # Сохраним результат заполнения логинов и паролей в файл
    $WorkBook.SaveAs($File)
    $WorkBook.Close($File)
    $Excel.Quit();
    }
}
)

#Выводим GUI
$window_form.Controls.Add($groupBox1)
$window_form.Controls.Add($groupBox2)
$window_form.Controls.Add($GoButton)
$window_form.Controls.Add($Or)


#чтобы принудительно открыть окно скрипта поверх других диалоговых окон
$window_form.Topmost = $true
$result = $window_form.ShowDialog()


