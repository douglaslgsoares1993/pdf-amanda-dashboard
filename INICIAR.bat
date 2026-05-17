@echo off
where python >nul 2>nul
if %errorlevel%==0 (
  python INICIAR.py
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 INICIAR.py
  ) else (
    echo Python nao encontrado. Execute INSTALAR.py primeiro.
  )
)
pause
