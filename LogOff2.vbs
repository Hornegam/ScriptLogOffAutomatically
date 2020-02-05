Set oWshShell = WScript.CreateObject("WScript.Shell")

'Seta o tempo para 6000000 que é 100 minutos
tempo = "6000000"

'Cria o objeto time
Dim time,x,pop


'Passa por parametro, o CDbl é um conversor de string para Integer
time = CDbl(tempo)

'Converte de Integer para Integer
min = CStr(time/60000) 

'Usa o min que foi convertido para poder escrever

pop = oWshShell.Popup("O computador ira desligar em "+min+" minutos!", 30)
pop = oWshShell.Popup("O computador ira desligar caso fique 10 minutos sem mexer !", 30)

'Comando para dormir o tempo estimado
WScript.Sleep time 'Timer is adjusted to suit your needs

sCommand = "logoff.exe"
oWshShell.Run(sCommand)