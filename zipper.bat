::pc only works with 7zip right now

if exist {C:\Program Files\7-Zip\7z.exe} (
    for %A in (*) do "C:\Program Files\7-Zip\7z.exe" a -tzip  %~nA.zip %A
) 
else (
	start /d "C:\Windows\System32\" compact.exe /u
)

