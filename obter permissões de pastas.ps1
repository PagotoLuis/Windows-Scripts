# Define o diret�rio a ser verificado e o caminho para o arquivo CSV de sa�da
$directory = "D:/"
$outputCsv = "C:/"

# Fun��o para obter as permiss�es de uma pasta
function Get-FolderPermissions {
    param (
        [string]$folderPath
    )
    
    $acl = Get-Acl -Path $folderPath
    $permissions = @()

    foreach ($access in $acl.Access) {
        $permissions += [PSCustomObject]@{
            FolderPath = $folderPath
            IdentityReference = $access.IdentityReference
            FileSystemRights = $access.FileSystemRights
            AccessControlType = $access.AccessControlType
        }
    }

    return $permissions
}

# Obt�m todas as pastas diretamente dentro do diret�rio especificado (sem subpastas)
$folders = Get-ChildItem -Path $directory -Directory

# Inicializa uma lista para armazenar todas as permiss�es
$allPermissions = @()

# Itera por cada pasta e obt�m suas permiss�es
foreach ($folder in $folders) {
    $folderPermissions = Get-FolderPermissions -folderPath $folder.FullName
    $allPermissions += $folderPermissions
}

# Exporta as permiss�es para um arquivo CSV
$allPermissions | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Output "Relat�rio de permiss�es gerado em $outputCsv"