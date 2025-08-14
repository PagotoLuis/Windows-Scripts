$status = wmic diskdrive get status

if ($status -like "*OK*"){
	"OK" | Out-File -FilePath "C:\Users\luis.pagoto\Desktop\teste.txt" -Append
} else {
	"ERRO" | Out-File -FilePath "C:\Users\luis.pagoto\Desktop\teste.txt" -Append
}