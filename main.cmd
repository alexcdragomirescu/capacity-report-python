@ECHO OFF

SET WD=%~dp0
SET WD=%WD:~0,-1%
SET ENTRY=%WD%\main.py

SET SCRIPTS=%WD%\.venv\Scripts
SET PY_EXEC=%SCRIPTS%\python.exe

CMD /K %PY_EXEC% %ENTRY%
