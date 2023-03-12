#Скрипт для рассылки через телеграмм-бота уведомлений сотрудникам.
$token = "52222222:111111111111"

# Telegram URLs
$URL_get = "https://api.telegram.org/bot$token/getUpdates"
$URL_set = "https://api.telegram.org/bot$token/sendMessage" 
#$Proxy = "http://192.168.10.20:80"

function sendMessage($URL, $chat_id, $text)
{
    # создаем HashTable, можно объявлять ее и таким способом
    $ht = @{
        text = $text
        # указан способ разметки Markdown
        parse_mode = "HTML"
        chat_id = $chat_id
    }
    # Данные нужно отправлять в формате json
    $json = $ht | ConvertTo-Json
    # Делаем через Invoke-RestMethod, но никто не запрещает сделать и через Invoke-WebRequest
    # Method Post - т.к. отправляем данные, по умолчанию Get
    #Invoke-RestMethod $URL -Method Post -ContentType 'application/json; charset=utf-8' -Body $json -Proxy $Proxy | Out-Null   
    Invoke-RestMethod $URL -Method Post -ContentType 'application/json; charset=utf-8' -Body $json | Out-Null 
}


# Подключение модуля для работы с ActiveDirectory
Import-Module ActiveDirectory

# Узнаем дату "за 5 ней до блокировки"
$lastLogontimestamp6m = (Get-Date).AddMonths(-6).AddDays(5).ToFileTime().ToString()

# Берем дату 3 месяца. Чтобы исключить свежесозданные учетки.
$Created3M = (Get-Date).AddMonths(-3).ToString()

#Делаем выборку пользователей. В определенной OU. Включенные, созданные более 3х месяцев, с датой lastLogon за 5 ней до блокировки, или которые не входили ни разу
$UserBan = get-aduser -Properties LastLogonDate, pager, lastLogontimestamp, Created  -Filter {(Enabled -eq "True") -and (Created -lt $Created3M )  -and ((lastLogontimestamp -le $lastLogontimestamp6m) -or (-not (lastlogontimestamp -like "*"))) } -SearchBase "OU=Domain Users,DC=contoso,DC=local"  | Select SamAccountName , lastLogontimestamp, LastLogonDate, Created, pager | ?{$_.pager -gt "1"}  | sort lastLogontimestamp
$UserBan | Select pager
$UserBan | ForEach-Object {
    $Chat_id = $_.pager

    #Собираем сообщение
    $msg = 'Добрый день, через 5 дней Ваша учетная запись будет заблокирована во внутренних сервисах Contoso и удалена из информационного Телеграм-канала из-за отсутствия активности учетной записи в течение полугода.' + "`n"
    $msg += "`n"
    $msg += 'Чтобы не было блокировки, рекомендуем пройти авторизацию в одном из сервисов:' + "`n"
    $msg +=  '- https://contoso.com' + "`n"
    $msg += "`n"
    $msg += 'При возникновении вопросов обращайтесь, пожалуйста, в службу технической поддержки'  
  
  
    Write-Host $Chat_id  "`n" $msg -f yellow 
    
    #Отправляем сообщение в бот
    sendMessage $URL_set $Chat_id "$msg"
    }