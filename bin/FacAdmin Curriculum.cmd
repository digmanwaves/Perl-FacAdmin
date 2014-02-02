@echo off
FOR /F "tokens=1-2*" %%A IN ('REG QUERY "HKLM\Software\Digital Manifold Waves\FacAdmin" /v Path') DO set FAPath=%%C
perl "%FAPath%\facAdmin.pl" --curriculum %*
set FAPath=
echo FacAdmin finished.
pause