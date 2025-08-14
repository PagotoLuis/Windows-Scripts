# Define o diretório a ser verificado e o caminho para o arquivo CSV de saída
$directory = "D:/"
$outputCsv = "C:/"

# Função para obter as permissões de uma pasta
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

# Obtém todas as pastas diretamente dentro do diretório especificado (sem subpastas)
$folders = Get-ChildItem -Path $directory -Directory

# Inicializa uma lista para armazenar todas as permissões
$allPermissions = @()

# Itera por cada pasta e obtém suas permissões
foreach ($folder in $folders) {
    $folderPermissions = Get-FolderPermissions -folderPath $folder.FullName
    $allPermissions += $folderPermissions
}

# Exporta as permissões para um arquivo CSV
$allPermissions | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Output "Relatório de permissões gerado em $outputCsv"