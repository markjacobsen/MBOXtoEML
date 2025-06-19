set CODE_BASE=%~dp0

xcopy "%CODE_BASE%bin\Debug\net8.0\" "%SYNC_DRIVE_HOME%\Apps\CFG2\MBOXtoEML\" /Y /E
