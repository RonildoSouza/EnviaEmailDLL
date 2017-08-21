@ECHO OFF
SET pathDotNetVersion4="%windir%\Microsoft.NET\Framework\v4.0.30319"

REG QUERY HKCR\EnviaEmailDLL.Configuracao
IF %ERRORLEVEL% EQU 1 (
    CLS
    IF EXIST %pathDotNetVersion4% (
        ECHO *** GERANDO ARQUIVO DE REGISTRO... ***
        "%pathDotNetVersion4%\RegAsm.exe" EnviaEmailDLL.dll /regfile:"RegistraDLL.reg"    
        COLOR A
        IF %ERRORLEVEL% EQU 1 COLOR C
    ) ELSE (
        COLOR C
        ECHO "dotNet Framework versao 4.5 nao instalada!"
    )
) ELSE (
    CLS
    COLOR C
    ECHO "A DLL ENVIAEMAILDLL.DLL JA ESTA REGISTRADA!"
)

PAUSE