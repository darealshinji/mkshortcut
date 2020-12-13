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
* Visual Studio 2019: open the provided solution file (mkshortcut.sln) and select "Build" -> "Build solution"
* MSVC command line: cl.exe /W2 /O2 mkshortcut.cpp
* GCC: x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -municode -o mkshortcut.exe  mkshortcut.cpp  -lole32 -luuid -static -s

