:loop
nircmd.exe changesysvolume 10000
nircmd.exe waitprocess chrome.exe speak text "You've got mail"
nircmd.exe changesysvolume -10000
ping localhost -n 922 >= nul
goto loop

