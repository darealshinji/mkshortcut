This is a simple C++ command line program to create so-called Shell Links, also known as Shortcuts or .lnk files.

Motivation for writing it was the lack of such a tool coming preinstalled on Windows 10 (mklink.exe
is for hard/symbolic links) and to get a better understanding on how to use the IShellLink class.
It was published into its own repository in the hope that people will have it easier to find it.
Check out https://github.com/libyal/liblnk if you don't want to use the Windows API.

Notes:
* released into the Public Domain (Unlicense)
* does not create .url files (those are text files in an INI format)
* no header file; just copy the code you need from mkshortcut.cpp if you want to use it in your project
* Unicode only; it's not the 90's anymore and the used IPersistFile class only takes LPCOLESTR, which is wide character
* requires linkage against `ole32.lib` on MSVC and `-lole32 -luuid` on GCC/MinGW

Compile:
* Visual Studio 2019: open the provided solution file (mkshortcut.sln) and select *Build* -> *Build solution*
* MSVC command line: `cl.exe /W2 /O2 mkshortcut.cpp`
* GCC: `x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -municode -o mkshortcut.exe  mkshortcut.cpp  -lole32 -luuid -static -s`


Usage example
-------------
*(assuming Firefox installed in typical directory)*

Open a command promt, move to Firefox's directory and create a shortcut:
```
cd "C:\Program Files\Mozilla Firefox"

C:\path\to\mkshortcut.exe /o:"%homedrive%%homepath%\Desktop\Firefox in disguise.lnk" ^
  /t:firefox.exe /tfull /w:"C:/" /a:Windows ^
  /i:"..\Internet Explorer\iexplore.exe" /ifull ^
  /d:"Firefox (really)" /k:saf
```

`/o:"%homedrive%%homepath%\Desktop\Firefox in disguise.lnk"`
 -> the shortcut will be placed on your desktop and called "Firefox in disguise" based on the filename

`/t:firefox.exe` -> shortcut target relative to your working directory

`/tfull` -> resolves relative path *firefox.exe* to the full path *"C:\Program Files\Mozilla Firefox\firefox.exe"*

`/w:"C:/"` -> set the working directory to *C:/*

`/a:Windows` -> launch Firefox with *Windows* as command line parameter;
this will make Firefox open *Windows* wich is resolved from the working directory (C:/) and therefor the contents of *C:\Windows* is shown

`/i:"..\Internet Explorer\iexplore.exe"` -> use the Internet Explorer icon

`/ifull` -> resolves *"..\Internet Explorer\iexplore.exe"* to the full path *"C:\Program Files\Internet Explorer\iexplore.exe"*

`/d:"Firefox (really)"` -> set tooltip description

`/k:saf` -> set "hotkey" to **S**hift+**A**lt+**F**; pressing this combination when being "on the desktop" will open the shortcut
