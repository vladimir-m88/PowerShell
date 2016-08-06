# Этот скрипт формирует файл с данными о пользователях, которые имеют доступ к коллекции сайтов (к сайту верхнего уровня)

# 2016_08_03 - проверен на http://mih-spsdev-03:1111, http://mih-spsdev-03:3333 - все ОК.
# 2016_08_04 - проверен на https://sgk.metinvest.ua, https://spp.metinvestholding.com - все ОК.

# Подключение модулей.
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Функция, которая возвращает название компанни, где работает пользователь.
function getCompany([string]$login)
{
    if ($login.Length -le 0) { return '-'}
    [string]$domainName = ($login.split("\")[0]) -replace "^.+\|", ""
    [string]$userName = $login.split("\")[1]
    Try
    {
        [Microsoft.ActiveDirectory.Management.ADUser]$user = Get-ADUser $userName -Server $domainName -Properties Company
        [string]$company = $user.Company
        if ($company.Length -gt 0) {
            return $company
        } else {
            return '-'
        }
    }
    Catch
    {
        Write-Host $('Exeption in function getCompany(): $domainName = ' + $domainName + '; $userName = ' + $userName + '; $_.Exception.Message: ' + $_.Exception.Message) -ForegroundColor Red
        return '-'
    }
}

# функция, которая формирует массив объектов с данными пользователей, входящих в группу AD.
function GetGroupMembers([string]$groupName,  [ref]$users)
{
    $group = Get-ADGroup $groupName -Properties Members
    foreach($dn in $group.Members)
    {
        [string]$Name = "-"
        if($dn.startswith("CN=S-1-5")) # пользователь из внешнего домена
        {
            [string]$SIDText = ($dn.Split(","))[0].SubString(3)
            $SID = New-Object System.Security.Principal.SecurityIdentifier $SIDText
            [string]$UserLogin = $($SID.Translate([System.Security.Principal.NTAccount]).Value)
            [string]$domainName = $UserLogin.split("\")[0]
            [string]$userName = $UserLogin.split("\")[1]
            try
            {
                $Name = (Get-ADUser $userName -Server $domainName).Name
            }
            catch
            {
                Write-Host $("Exeption in function GetGroupMembers(), getting user from a trust domain: " + $_.Exception.Message) -ForegroundColor Red
            }
        }
        else # пользователь или группа из текущего домена
        {
            Try
            {
                # Если $dn указывает на группу, то попадем в Catch
                $ADUser = get-aduser $dn -properties cn,samaccountname,EmailAddress
                $UserLogin = "metinvest\$($ADUser.SamAccountName)"
                $Name = $ADUser.Name
            }
            Catch
            {
                Try
                {
                    # Считаем что $dn указывает на группу
                    $subgroup = Get-ADGroup $dn
                    GetGroupMembers -groupName $subgroup.Name -users $users # Рекурсивный вызов
                }
                Catch
                {
                    Write-Host $("Exeption in function GetGroupMembers(), in Get-ADGroup or recursive call: " + $_.Exception.Message) -ForegroundColor Red
                }
            }
        }
        $tmp = @{
            AdGroup = $group.Name
            UserLogin = $UserLogin
            Name = $Name
        }
        $user = new-object psobject -Property $tmp
        $users.Value += $user
    }
}

# Функция, которая возвращает строку разрешений пользователя.
function GetPermissions([Microsoft.SharePoint.SPRoleCollection]$roles)
{
    [string]$rolesStr = ""
    if ($roles.Length -ge 1) {
        foreach ($role in $roles) {
            if (($role.ToString() -ne "Limited Access") -and ($role.ToString() -ne "Ограниченный доступ")) {
                $rolesStr += $role.ToString() + "; "
            }
        } 
    }
    return $rolesStr
}

# Функция, которая записывает данные пользователя в файл.
function LogUserData ([Microsoft.SharePoint.SPUser]$user, [Microsoft.SharePoint.SPRoleCollection]$roles, [string]$FileUrl) 
{
    [string]$rolesStr = GetPermissions -roles $roles
    if ($rolesStr.Length -gt 0) {
        if ($user.IsDomainGroup -eq $False)
        {
            "права предоставлены не через AD-группу `t $($user.UserLogin) `t $($user.Name) `t $(getCompany -login $user.UserLogin) `t $($rolesStr)" | out-file $FileUrl -Append
        }
        else
        {
            [string]$groupName = $user.Name.split("\")[1]
            if (($groupName.length -gt 0) -and ($groupName -ne "authenticated users"))
            {
                $adusers = @()
                GetGroupMembers -groupName $groupName -users  ([ref]$adusers)
                foreach ($aduser in $adusers)
                {
                    "$($aduser.AdGroup) `t $($aduser.UserLogin) `t $($aduser.Name) `t $(getCompany -login $aduser.UserLogin) `t $($rolesStr)" | out-file $FileUrl -Append
                }
            }
        }
    }
}

# Основная функция, формирующая файл с данными о пользователях, которые имеют доступ к коллекции сайтов (к сайту верхнего уровня).
function GetListOfAllUsers ([string]$SiteUrl, [string]$FileUrl)
{
    "Доменная группа `t Учетная запись `t Имя пользователя `t Предприятие `t Разрешения" | out-file $FileUrl
    $site = Get-SPSite $SiteUrl
    # Пользователи, доступ которым предоставлен непосредственно.
    #$site.RootWeb.SiteUsers[0] | Get-Member
    foreach ($user in $site.RootWeb.SiteUsers) {
        LogUserData -user $user -roles $user.Roles -FileUrl $FileUrl
    }

    # Пользователи, доступ которым предоставлен через группы SharePoint.
    $groups = $site.RootWeb.SiteGroups
    foreach ($grp in $groups)
    {
        if ($grp.Name -eq "Читатели ресурсов стилей") {continue}
        foreach ($user in $grp.Users)
        {
            LogUserData -user $user -roles $grp.Roles -FileUrl $FileUrl
        }
    }
}

#
GetListOfAllUsers -SiteUrl "https://spp.metinvestholding.com" -FileUrl "c:\tmp\SiteCollectionUsers.csv" #http://mih-spsdev-03:3333; https://sgk.metinvest.ua