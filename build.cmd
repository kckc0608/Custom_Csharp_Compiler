:: MSBuild 를 호출하여 빌드하는 배치파일.
:: echo off 로 하단에 작성한 명령어의 배치 프로그램상 출력을 끈다.
:: 명령어를 실행하고, 배치프로그램이 바로 종료되지 않도록 Pause 명령어로 일시중지 시킨다.
@echo off
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild WPF_IDE.csproj
Pause