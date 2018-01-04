:loop
nircmd.exe changesysvolume 10000
nircmd.exe speak text "Classic Kyle!"
nircmd.exe changesysvolume -10000
ping localhost -n 922 >= nul
goto loop

