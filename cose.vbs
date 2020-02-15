do
a = Inputbox("cosa vuoi che dica?","voce","inserire testo")
if a = "sono una cretina" Or a = "sono una troia" Or a = "sono stupida" Or a = "sono una sgualdrina" then
CreateObject("SAPI.spVoice").speak "vaffanculo"
elseIf a = "" then
wscript.quit
else
CreateObject("SAPI.spVoice").speak a
end if
loop