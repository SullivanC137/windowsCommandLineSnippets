# use this to search for and replace a word in all files in current folder
# Following lines check all files with extension .ins and file name starting with sto_vld
# and then replaces LAATSTE_VISREIS_VERSIE_IND with LAATSTE_VISREIS_VERSIE_INDIC

gci -r -include "sto_vld_*.ins" |
foreach-object { $a = $_.fullname; ( get-content $a ) |
foreach-object { $_ -replace "LAATSTE_VISREIS_VERSIE_IND","LAATSTE_VISREIS_VERSIE_INDIC" }  |
set-content $a }

# alternatively, use Bing chat, powered by chatgpt
# prompt: write a powershell script that replaces a string in all files in a folder
# following code is NOT tested
$Path = "C:\Folder\*.*"
$OldString = "OldString"
$NewString = "NewString"

Get-ChildItem $Path -Recurse | ForEach-Object {
    (Get-Content $_.FullName) |
    ForEach-Object { $_ -replace $OldString, $NewString } |
    Set-Content $_.FullName
}
