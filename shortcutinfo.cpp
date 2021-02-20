/*
 * This is free and unencumbered software released into the public domain.
 *
 * Anyone is free to copy, modify, publish, use, compile, sell, or
 * distribute this software, either in source code form or as a compiled
 * binary, for any purpose, commercial or non-commercial, and by any
 * means.
 *
 * In jurisdictions that recognize copyright laws, the author or authors
 * of this software dedicate any and all copyright interest in the
 * software to the public domain. We make this dedication for the benefit
 * of the public at large and to the detriment of our heirs and
 * successors. We intend this dedication to be an overt act of
 * relinquishment in perpetuity of all present and future rights to this
 * software under copyright law.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
 * OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 *
 * For more information, please refer to <http://unlicense.org/>
 */

/**
 * Compile with GCC/MinGW:
 *   x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -municode -o shortcutinfo.exe  shortcutinfo.cpp  -lole32 -luuid -static -s
 *
 * Compile with MSVC:
 *   cl.exe /W3 /O2 shortcutinfo.cpp
 */

// TODO: show CLSID targets

#ifdef _MSC_VER
# pragma comment(lib, "ole32.lib")
#endif
#ifndef UNICODE
# define UNICODE
#endif
#include <windows.h>
#include <objbase.h>
#include <shlobj.h>
#include <stdio.h>
#include <wchar.h>


int wmain(int argc, wchar_t *argv[])
{
	IShellLink *shLink = NULL;
	IPersistFile *pFile = NULL;
	HRESULT hRes;
	int ret = 1;

	wchar_t buf[4096];
	const int bufLen = ARRAYSIZE(buf);
	LPWSTR pbuf = reinterpret_cast<LPWSTR>(&buf);
	WORD wHotkey = 0;
	int n = 0;

	if (argc < 2) {
		wprintf_s(
			L"Shows information about Shell Links\n"
			L"usage: %s FILENAME\n", argv[0]);
		return 0;
	}

	// initialize COM library
	hRes = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE | COINIT_SPEED_OVER_MEMORY);

	if (hRes != S_OK) {
		wprintf_s(L"%s: could not initialize COM library\n", argv[0]);
		return 1;
	}

	// create instance
	hRes = CoCreateInstance(
		CLSID_ShellLink,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IShellLink,
		reinterpret_cast<void **>(&shLink));

	if (hRes != S_OK) {
		goto MAIN_EXIT;
	}

	// query interface
	hRes = shLink->QueryInterface(IID_IPersistFile, reinterpret_cast<void **>(&pFile));

	if (hRes != S_OK) {
		goto MAIN_EXIT;
	}

	// open file
	if (pFile->Load(argv[1], 0) != S_OK) {
		goto MAIN_EXIT;
	}

	// path
	if (shLink->GetPath(pbuf, bufLen, NULL, 0) == S_OK && wcslen(buf) > 0) {
		wprintf_s(L"Target path: %s\n", buf);
	}

	// arguments
	if (shLink->GetArguments(pbuf, bufLen) == S_OK && wcslen(buf) > 0) {
		wprintf_s(L"Arguments: %s\n", buf);
	}

	// description
	if (shLink->GetDescription(pbuf, bufLen) == S_OK && wcslen(buf) > 0) {
		wprintf_s(L"Description: %s\n", buf);
	}

	// icon
	if (shLink->GetIconLocation(pbuf, bufLen, &n) == S_OK && wcslen(buf) > 0) {
		wprintf_s(
			L"Icon location: %s\n"
			L"Icon index: %d\n", buf, n);
	}

	// working directory
	if (shLink->GetWorkingDirectory(pbuf, bufLen) == S_OK && wcslen(buf) > 0) {
		wprintf_s(L"Working directory: %s\n", buf);
	}

	// showcmd
	if (shLink->GetShowCmd(&n) == S_OK) {
		wprintf_s(L"Show command: ");

		switch(n) {
			case SW_SHOWNORMAL:
				wprintf_s(L"SW_SHOWNORMAL\n");
				break;
			case SW_SHOWMAXIMIZED:
				wprintf_s(L"SW_SHOWMAXIMIZED\n");
				break;
			case SW_SHOWMINIMIZED:
				wprintf_s(L"SW_SHOWMINIMIZED\n");
				break;
			default:
				wprintf_s(L"0x%x (not supported)\n", n);
				break;
		}
	}

	// hotkey
	if (shLink->GetHotkey(&wHotkey) == S_OK && wHotkey != 0) {
		wprintf_s(L"Hotkey:");

		unsigned char hi = (wHotkey >> 8) & 0xff;
		unsigned char lo = wHotkey & 0xff;

		if (hi & HOTKEYF_CONTROL) {
			wprintf_s(L" CTRL");
		}

		if (hi & HOTKEYF_SHIFT) {
			wprintf_s(L" SHIFT");
		}

		if (hi & HOTKEYF_ALT) {
			wprintf_s(L" ALT");
		}

		if (hi & HOTKEYF_EXT) {
			wprintf_s(L" EXT");
		}

#ifdef _MSC_VER
		if ((lo >= 'A' && lo <= 'Z') || (lo >= '0' && lo <= '9')) {
			wprintf_s(L" %c", lo);  // A-Z, 0-9
		} else if (lo >= 0x70 && lo <= 0x7b) {
			wprintf_s(L" F%d", lo - 0x6f);  // F1 ... F12
		} else if (isgraph(lo)) {
			wprintf_s(L" %c", lo);
		} else {
			wprintf_s(L" 0x%x", lo);
		}
#else
		switch(lo) {
			case 'A' ... 'Z':
			case '0' ... '9':
				wprintf_s(L" %c", lo);
				break;
			case 0x70 ... 0x7b:  // F1 ... F12
				wprintf_s(L" F%d", lo - 0x6f);
				break;
			default:
				if (isgraph(lo)) {
					wprintf_s(L" %c", lo);
				} else {
					wprintf_s(L" 0x%x", lo);
				}
				break;
		}
#endif

		wprintf_s(L" (0x%x)\n", wHotkey);
	}

	ret = 0;

MAIN_EXIT:

	if (shLink) {
		shLink->Release();
	}

	if (pFile) {
		pFile->Release();
	}

	CoUninitialize();

	return ret;
}
