#importa os modulos para acessar o exchange online
Import-Module ExchangeOnlineManagement

$usuario = Read-Host "Usuario: "
$nome = Read-Host "Nome MAIUSCULO: "
$dominio = Read-Host "dominio: "

#Faz o donwload das informações da caixa
$formatenumerationlimit = -1
Get-Mailbox "$usuario@metasa.com.br" | fl > "$usuario.txt"

#Desativa a caixa local
Disable-Mailbox "$dominio@metasa.com.br"

#Habilite o usuário local como uma caixa de correio remota:
Enable-RemoteMailbox "$usuario@$dominio.com.br" -RemoteRoutingAddress "$usuario@$dominio.mail.onmicrosoft.com"

#Adicione o LegacyExchangeDNvalor da caixa de correio local anterior ao endereço de proxy da nova caixa de correio remota como um endereço x500. Para isso, execute o seguinte cmdlet:
Set-RemoteMailbox -Identity "$usuario@$dominio.com.br" -EmailAddresses @{add="x500:/o=Exchange $dominio/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=<$nome>"}

#Conectar no exchange online para pegar o guid do caixa na nuvem
Connect-ExchangeOnline -UserPrincipalName "$dominio@$dominio.com.br"
Get-Mailbox "$usuario@$dominio.com.br" | fl *ExchangeGUID*

#Pegar o GUID da caixa local
Get-MailboxDatabase "Perfil2" | fl *GUID*

$cred = Get-Credential
#New-MailboxRestoreRequest -RemoteHostName "mail.contoso.com" -RemoteCredential $cred -SourceStoreMailbox "exchange guid of disconnected mailbox" -TargetMailbox "exchange guid of cloud mailbox" -RemoteDatabaseGuid "guid of on-premises database" -RemoteRestoreType DisconnectedMailbox
New-MailboxRestoreRequest -RemoteHostName "UMA052S.metasa.com.br" -RemoteCredential $cred -SourceStoreMailbox "75c61781-3b6e-4cd3-bdea-f70e9811d460" -TargetMailbox "86a70554-bbc7-4297-8f0c-c21e563ec79c" -RemoteDatabaseGuid "85011388-271f-48d3-addd-35eebe3a2cc0" -RemoteRestoreType DisconnectedMailbox
