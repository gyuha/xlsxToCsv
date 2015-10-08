@rem exe파일에 dll을 합치기
@rem Download http://www.microsoft.com/en-us/download/confirmation.aspx?id=17630
"C:\Program Files (x86)\Microsoft\ILMerge\ILMerge.exe" bin\release\xlsxToCsv.exe Excel.dll ICSharpCode.SharpZipLib.dll /out:xlsxToCsv.exe /lib:"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0" /targetplatform:v4,"C:\Windows\Microsoft.NET\Framework64\v4.0.30319"