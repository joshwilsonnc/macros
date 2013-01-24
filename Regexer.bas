Attribute VB_Name = "Regexer"
Sub regEx()
Attribute regEx.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Regex Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    
    'Create regexp object
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")
    'Dim regEx As New VBScript_RegExp_55.RegExp
        
    Dim toReplace As Boolean
    Dim str As String
    Dim find As String
        
    'Get find and replace patters from regexForm
    regexForm.Show
    find = regexForm.regexFind.Value
    
    If find <> "" Then
    
        'set find pattern
        regEx.Pattern = regexForm.regexFind.Value
        
        'initialize replacement with empty string
        str = ""
        
        'set replacement string, if requested
        If regexForm.regexReplacement.Value = True Then
            toReplace = True
            str = regexForm.regexReplace.Value
        Else
            toReplace = False
        End If
        
        'set match options
        If regexForm.regexCase.Value = True Then
            regEx.IgnoreCase = True
        End If
        If regexForm.regexGlobal.Value = True Then
            regEx.Global = True
        End If
        If regexForm.regexMulti.Value = True Then
            regEx.MultiLine = True
        End If
        
        If toReplace Then
            Do While (Not IsEmpty(ActiveCell.Value))
                ActiveCell.Value = regEx.Replace(ActiveCell.Value, str)
                ActiveCell.Offset(1).Select
            Loop
        Else
            Do While (Not regEx.Test(ActiveCell.Value)) And (Not IsEmpty(ActiveCell.Value))
                ActiveCell.Offset(1).Select
            Loop
        End If
        
    End If
    
    Unload regexForm
    
End Sub
Function removeFirstLI(str)

    'removes first <li>See also [whatever]</li>
    
    Dim regEx As New VBScript_RegExp_55.RegExp
    
    regEx.Pattern = "\<li\>\s?See\s?also.+?\</li\>"
    regEx.IgnoreCase = True 'True to ignore case
    regEx.Global = False 'True matches all occurences, False matches the first occurence
    
    removeFirstLI = regEx.Replace(str, "")
    
End Function

Function replaceFirstLI(str)

    'replaces first </li> with .@@
    
    Dim regEx As New VBScript_RegExp_55.RegExp
    
    regEx.Pattern = "\</li\>"
    regEx.IgnoreCase = True 'True to ignore case
    regEx.Global = False 'True matches all occurences, False matches the first occurence
    
    replaceFirstLI = regEx.Replace(str, "   .@@")
    
End Function

