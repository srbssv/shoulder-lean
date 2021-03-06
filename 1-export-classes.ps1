Import-Module ActiveDirectory

# Экспортировать несколько классов и параллелей в csv

$ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$Classes_ = Read-Host("Введи классы или параллели через пробел")
$Classes = $Classes_.Split(" ")
$ClassFinal = @()
    # Разбираем абракадабру
    foreach ($class in $Classes){
        # Если ввели параллель (только число), тогда вытаскиваем все классы по этой параллели
        if ($class -match "^\d+$") {
            $class_ = $class + "*"
            $OU_ = Get-ADOrganizationalUnit -Filter {Name -like $class_} -SearchBase "OU=Students, OU=Users, OU=ALM-FM, dc=nis, dc=edu, dc=kz" -SearchScope OneLevel | Select Name
            foreach ($OU__ in $OU_) {
                $ClassFinal += $OU__.Name.ToString()
            }
        }
        # Если ввели название класса, то вытаскиваем только класс
        else {
           $ClassFinal = $ClassFinal + $class
        }
    }
    # Теперь вытаскиваем согласно списку ClassFinal инфу по ученикам из ActiveDirectory и сохраняем их в csv
foreach ($class in $ClassFinal) {
    $OU = "OU="+$class+", OU=Students, OU=Users, OU=ALM-FM, dc=nis, dc=edu, dc=kz"
    $CSVPath = $ScriptPath + "\" + $class + ".csv"
    Get-ADUser -Filter * -SearchBase $OU -SearchScope OneLevel | Select GivenName, Surname, SamAccountName, UserPrincipalName | Export-CSV $CSVPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
}
Read-Host("Все норм, нажми Enter")
# Пьем чай