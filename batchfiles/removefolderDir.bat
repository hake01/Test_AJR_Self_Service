del /q "D:\sa_AJRService1\workspace\AJR_test\DownloadingTempfiles\*"
FOR /D %%p IN ("D:\sa_AJRService1\workspace\AJR_test\DownloadingTempfiles\*.*") DO rmdir "%%p" /s /q