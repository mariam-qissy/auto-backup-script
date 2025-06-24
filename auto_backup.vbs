Option Explicit

Dim sourceFolder, backupFolder, dateStr, destination
Dim fso

' تحديد مجلد الملفات الأصلية
sourceFolder = "C:\Users\maria\Documents\my_files"

' تحديد مجلد النسخ الاحتياطي
backupFolder = "C:\Users\maria\Documents\backup"

' توليد التاريخ الحالي لتسمية النسخة
dateStr = Replace(Replace(Now, ":", "-"), " ", "_")
destination = backupFolder & "\backup_" & dateStr

Set fso = CreateObject("Scripting.FileSystemObject")

' إنشاء مجلد النسخة إذا ما كان موجود
If Not fso.FolderExists(backupFolder) Then
    fso.CreateFolder(backupFolder)
End If

' إنشاء مجلد جديد داخل مجلد النسخ
If Not fso.FolderExists(destination) Then
    fso.CreateFolder(destination)
End If

' نسخ كل الملفات والمجلدات
fso.CopyFolder sourceFolder & "\*", destination

WScript.Echo "✅ Backup created successfully at: " & destination
