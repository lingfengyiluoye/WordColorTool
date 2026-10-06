@echo off
cd /d %~dp0
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo /target:winexe /optimize+ /platform:x86 /out:WordColorTool.exe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll Program.cs MainForm.cs
if %errorlevel%==0 (echo BUILD OK) else (echo BUILD FAILED)
pause
