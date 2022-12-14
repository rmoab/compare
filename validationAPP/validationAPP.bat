@echo off
cls
echo -------------------------
echo        ValidationAPP
echo -------------------------
set PARAMETER1=
set /P PARAMETER1= Informe a localizacao dos arquivos XML: %=%
set PARAMETER2=
set /P PARAMETER2= Informe a localizacao do arquivo validador: %=%
java -jar "validation.jar" %PARAMETER1% %PARAMETER2%
set /p option=Fim. Presione uma tecla para finalizar