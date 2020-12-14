function Translit
{
    param([string]$inString)

    #хэш-таблица соответствия русских и латинских символов
    $Translit = 
    @{
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
        [char]'Е' = "E"
        [char]'ё' = "e"
        [char]'Ё' = "E"
        [char]'ж' = "zh"
        [char]'Ж' = "ZH"
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
        [char]'Х' = "KH"
        [char]'ц' = "ts"
        [char]'Ц' = "TS"
        [char]'ч' = "ch"
        [char]'Ч' = "CH"
        [char]'ш' = "sh"
        [char]'Ш' = "SH"
        [char]'щ' = "shch"
        [char]'Щ' = "SHCH"
        [char]'ъ' = ""
        [char]'Ъ' = ""
        [char]'ы' = "y"
        [char]'Ы' = "Y"
        [char]'ь' = ""
        [char]'Ь' = ""
        [char]'э' = "e"
        [char]'Э' = "E"
        [char]'ю' = "yu"
        [char]'Ю' = "YU"
        [char]'я' = "ya"
        [char]'Я' = "YA"
        [char]' ' = " "
    }

    $outString = "";
    $chars = $inString.ToCharArray();
    foreach ($char in $chars) {$outString += $Translit[$char]}
    return $outString;
}



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Создать компьютер в АД'
$form.Size = New-Object System.Drawing.Size(350,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Введите ФИО сотрудника:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $name = $textBox.Text
    if ($name)
    {
        # Находим всех пользователей по ФИО, сортируем по дате создания(по убыванию)
        $department = (Get-ADUser -Filter {Name -eq $name} -Properties * | select CanonicalName, Created | Sort-Object -Property Created -Descending)
        if ($department)
        {
            # Определяем департамент пользователя который создан последним
            $department = $department[0].CanonicalName.Split("/")[3]
            # Получаем описание департамента
            $description_department = (Get-ADOrganizationalUnit -Filter {Name -eq $department} -Properties *).Description
            
            if (($description_department.Length -ge 15) -or ($null -eq $description_department))
            {
                Write-Warning "Выполненние скрипта прервано с ошибкой: "
                Write-Warning "Описание департамента слишком длинное либо пустое: $description_department"
            }
            else
            {
                
                $search_computer_name_mask = -join("MSK-",$description_department)
                $search_computer_name_mask = -join($search_computer_name_mask, "0*")
                
                $get_adcomputer = (Get-ADComputer -Filter {Name -like $search_computer_name_mask} -Properties * | select Name, Created | Sort-Object -Property Created -Descending)
                if ($get_adcomputer)
                {
                    # получаем 2 полседних компьютера
                    
                    $search_computer_name_mask = $search_computer_name_mask.Replace("*", "")
                    $first_computer = $get_adcomputer[0].Name
                    if ($get_adcomputer.Count -gt 1)
                    {
                        $second_computer = $get_adcomputer[1].Name
                    }
                    else
                    {
                        $second_computer = $search_computer_name_mask + "0"
                    }

                    $first = ($first_computer.Replace($search_computer_name_mask, "")) -as [int]
                    $second = ($second_computer.Replace($search_computer_name_mask, "")) -as [int]

                    # проверяем номера двух последних компьютеров
                    if ($first -gt $second)
                    {
                        $first = ($first + 1) -as [String]
                        $new_computer_name = $first_computer.Remove($first_computer.Length - $first.Length) + $first
                    }
                    else
                    {
                        $second = ($second + 1) -as [String]
                        $new_computer_name = $second_computer.Remove($second_computer.Length - $second.Length) + $second
                    }
                }
                else
                {
                    $search_computer_name_mask = $search_computer_name_mask.Replace("*", "")
                    $new_computer_name = $search_computer_name_mask

                    for ($i=1; $i -le (14 - $search_computer_name_mask.Length); $i++)
                    {
                        $new_computer_name += "0"
                    }
                    $new_computer_name += "1"
                }

            $name_translit = Translit($textBox.Text);
            $ldap_path = "указываем путь до OU"
            New-ADComputer -Name $new_computer_name -Description $name_translit -Path $ldap_path

            echo ""
            echo ""
            echo "Имя нового компьютера: $new_computer_name"
            echo ""
            echo ""
            }
        }
        else
        {
            Write-Warning "Выполненние скрипта прервано с ошибкой: Пользователь $name не найден"
        }
    }
    else
    {
        Write-Warning "Выполненние скрипта прервано с ошибкой: Введите ФИО сотрудника"
    }
}
Remove-Variable -Name * -Force -ErrorAction SilentlyContinue
powershell.exe

