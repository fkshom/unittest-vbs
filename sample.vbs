Call Include("base.vbs")
Call Include("stringbuilder.class.vbs")
Call Include("unittest.class.vbs")

Sub Include(ByVal strFile)
    Dim objFSO, objStream, strPath, strDir
    
    '���g�̃t�@�C���p�X���擾
    strPath = WScript.ScriptFullName
    
    '�e�f�B���N�g�����擾
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    strDir = objFSO.GetFile(strPath).ParentFolder
    
    '�w�肵��VBScript�t�@�C����ǂݍ���
    Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1, False)
    ExecuteGlobal objStream.ReadAll()
    objStream.Close
    
    Set objStream = Nothing
    Set objFSO = Nothing
End Sub

function factorial(n)
    if(vartype(n) <> 2) then factorial = (-1) : exit function
    if(n < 0) then factorial = (-1) : exit function
    if(n = 0) then factorial = 1 : exit function
    factorial = n * factorial(n - 1)
end function

dim Tester : set Tester = new UnitTest

WScript.echo "Testing" & vbNewline
WScript.echo "=======" & vbNewline
WScript.echo Tester.test("factorial", array(-1), -1) & vbNewline
WScript.echo Tester.test("factorial", array(0), 1) & vbNewline
WScript.echo Tester.test("factorial", array(1), 1) & vbNewline
WScript.echo Tester.test("factorial", array(2), 2) & vbNewline
WScript.echo Tester.test("factorial", array(3), 6) & vbNewline
WScript.echo Tester.test("factorial", array("string"), -1) & vbNewline
WScript.echo Tester.test("factorial", array(2.718281828), -1) & vbNewline

WScript.echo vbNewline

set Tester = nothing

