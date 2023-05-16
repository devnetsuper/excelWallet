@echo off
:loop
python -c "import terminalWallet; terminalWallet.run()"
echo.
echo Press any key to run the program again or CTRL+C to exit...
pause >nul
goto loop