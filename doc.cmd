@echo off

echo ensuring dependencies...
.nuget\nuget.exe install packages.config -OutputDirectory packages -ExcludeVersion

echo assumes f# interactive is on your path...

echo fsi Doc.fsx...
fsi.exe Doc.fsx

echo copy documentation to 'gh-pages' subfolder...
xcopy "content" "gh-pages\content\" /s /y /u
xcopy "ShrinkRay.html" "gh-pages\index.html" /y

echo documentation complete!
echo view ./index.html to review :)