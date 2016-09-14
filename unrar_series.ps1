# This script is used for extracting multiple files from their respective
# rar archives and put them in their folders. It automatically finds WinRaR
# for you and stops if you don't have it...! It has a neat feature of keeping
# the files during the process or if you don't have space and or want them
# remove it each time it extracts one! :)
#
#
#
# NOTE: You need to change you privileges in powershell to allow this
#             script to run... by typing the following command:
#
#             PS>  Set-ExecutionPolicy Unrestricted
#
#             --> DON'T RUN THIS AS ADMINISTRATOR we want to set this for this
#                 session only... don't set it as the default local policy, which
#                 is what running powershell as administrator and typing this does!
#
#
# Oh... almost forgot... this is licensed under the MIT License, so do what the
# fuck you want with it... should you like it and buy me a beer I'd be extremely 
# happy :)
#
# Author: Andrew Grammenos
#
# Finally...Flames (...Comments/Suggestions) should be directed to my mail which I
# reluctantly give: andreas.grammenos@gmail.com
#
#
# :wq

# Let's create some variables
[int]$fcount = 0;            # holds the folder number
[string]$path = "";          # holds the parent path
[string]$initpath = $pwd;    # holds the initial path so we can go back
[string]$winrarpath = "";    # winrar path (for UnRaR.exe)
[bool]$deletefiles = $true;  # we delete the files when we extract
[bool]$movefiles = $false;   # we move the files to parent and delete with similar name

# assign it to the file
if ($args.Count -ne 1)
{ echo "`nUsage is ./script path, try again!"; exit; }
# check if the path is valid
if ((Test-Path -path $($args.get(0))) -ne $True) 
{ echo "`nPath doesn't seem to exist... try again with a working path!"; exit; }
else { $path = $(Resolve-Path $($args.Get(0))); }
# check if the path contains files
[int]$count = (Get-ChildItem $path | ? { $_.psiscontainer } | select fullname).Count;
if ($count -eq 0) 
{ echo "`nPath doesn't seem to contain any folders...what should I do? Hmm EXIT!"; exit; }

# now check if we have winrar!
if (((Test-Path -path 'C:\Program Files\WinRAR') -eq $True))
{ $winrarpath = 'C:\Program Files\WinRAR\UnRAR.exe'; }
elseif (((Test-Path -Path 'C:\Program Files (x86)\WinRaR') -eq $True))
{ $winrarpath = 'C:\Program Files (x86)\WinRaR\UnRaR.exe'; }
else 
{ echo "`nWinRar 32bit or 64bit doesn't seem to be present...grab it at rarlab.com until then exiting..."; exit; }

# ech winrar location
echo "`nWinRar location has been found an unrar executable is located in: `n";
echo " $winrarpath";

# ask the user if he wants to keep the files
$key = Read-Host "`nDo you want to keep the compressed rar files after extraction? [y/n] ";

if($key -ieq 'Y') 
{ $deletefiles = $false; echo "`nDeleting files atfer extraction disabled!"; }
else { echo "`nDeleting files after extraction is enabled!`n"; }

# ask the user if he wants to move the files up in parent and delete the
# folders that contained the archives (so you get only the video files extracted
# in parent directory!

# WARNING IT DELETES ALL FOLDERS IN THE FILE BECAUSE IT ASSUMES IT ONLY CONTAINS THE SERIES ONLY!
echo "`nIMPORTANT: IT DELETES ALL FOLDERS IN THE FILE BECAUSE IT ASSUMES IT ONLY 
CONTAINS THE SERIES ONLY!";
$key = Read-Host "`nDo you want to move extracted files in main folder and delete the 
folders that contained each archive? [y/n] ";

if($key -ieq 'Y')
{ $movefiles = $true; echo "`nEnabling moving folders to parent and delete folders!!"; }
else { echo "`nKeeping folders and files will be inside the folder they got extracted!"; }

# Path for extraction

# inform the user of what we are about to do
echo "`nWe got this directory as argument: $($args.Get(0))";
echo "The Resolved absolute path is the following: $path";
echo "`nWe will traverse in the following folders: ";
Get-ChildItem $path -r | ? { $_.psicontainer } | select fullname;

[string]$curpath = $path;
[string]$rarvar = "";
# now go into each of the files and well...unrar! 
for($i = 0; $i -lt $count; $i++) { 
    $curpath = $path + "\$((Get-ChildItem $path | ?{ $_.PSIsContainer } | Select-Object Name | Select -ExpandProperty "Name").get($i))";
    echo "Delving into: $curpath with $i";
    cd $curpath;
    # now get the name
    #echo $pwd;
	echo "`nPath is: $pwd`n";
    $rarvar = ls $curpath -name | Select-String -pattern ".rar";
	# check if it is an array or not to accommodate both numbering sytaxes
	# syntax 1: .rar, .r00, r01 etc
	if((ls $curpath -name | Select-String -pattern ".rar") -is [system.array])
	# syntax 2: part-1.rar, part-2.rar etc
		# get the first array element
		{$rarvar = (ls $curpath -name | Select-String -pattern ".rar")[0].Line} 
		# we don't need a case for the non-array
	
    if($rarvar -ieq "")
        { echo "`nNo rar's found in this directory moving on"continue;}
    else
        { echo "`nValid compressed file found: $rarvar`n"; 
          echo "`nPassing to WinRaR for extraction...!`n";
          # now extract
          &($winrarpath) e $rarvar;		  
    }

    
    # check for error
    if( $? -eq $False )
        { echo "`nWarning winrar failed extracting the following file:";
          echo "`n$curpath"; echo "`nContinuing"; } 
    else {
        if($deletefiles)
          { echo "`nNow deleting everything except what I have extracted and exceptions :)";
            Start-Sleep -Seconds 3; # pause to release the other one!
            # This deletes everything except the exceptions
            Remove-Item "$curpath\*" -Recurse -Exclude *.mkv,*.avi,*.srt,*.wmv; }
        if($movefiles)
          { echo "`nNow moving the extracted files in the parent folder";
            Move-Item "$curpath\*.mkv" "..\"; 
            Move-Item "$curpath\*.avi" "..\"; 
            Move-Item "$curpath\*.srt" "..\";
            Move-Item "$curpath\*.wmv" "..\"; }
    }
    # go back 
    cd ..;
    #echo $pwd;
    # ls
    
    echo "`n Passing to the next one!";
}

# if we set to remove the folders delete them!
if($movefiles)
  { echo "`nNow deleting folders that contained the archives as well";
    Get-ChildItem $path -r | ? { $_.psicontainer } | select fullname | Remove-Item -Recurse; }

echo "`nEverything has gone as planned! exiting gracefully!";
# change path to the initial one
cd $initpath;
# now exit
exit;
