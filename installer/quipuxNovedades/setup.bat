cd ocx
copy COMDLG32.OCX C:\Windows\System32
copy COMDLG32.OCX C:\Windows\SysWOW64
cd C:\Windows\SysWOW64
Regsvr32 COMDLG32.OCX

echo "Done!"