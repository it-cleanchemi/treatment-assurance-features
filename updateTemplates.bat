@ECHO OFF

FOR /F "tokens=1,2 delims=:" %%A in (ScriptIDList.txt) do (
    ECHO %%A
    
    @REM Make .clasp.json file
    (
        ECHO {
        ECHO   "scriptId": "%%B",
        ECHO   "rootDir": "GAS"
        ECHO }
    )> .clasp.json

    CALL clasp push --force
)

ECHO Copying Source File
COPY source.clasp.json .clasp.json
