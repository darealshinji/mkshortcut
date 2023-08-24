/*
Licensed under the MIT License <http://opensource.org/licenses/MIT>.
SPDX-License-Identifier: MIT
Copyright (c) 2020-2023 djcj@gmx.de

Permission is hereby  granted, free of charge, to any  person obtaining a copy
of this software and associated  documentation files (the "Software"), to deal
in the Software  without restriction, including without  limitation the rights
to  use, copy,  modify, merge,  publish, distribute,  sublicense, and/or  sell
copies  of  the Software,  and  to  permit persons  to  whom  the Software  is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE  IS PROVIDED "AS  IS", WITHOUT WARRANTY  OF ANY KIND,  EXPRESS OR
IMPLIED,  INCLUDING BUT  NOT  LIMITED TO  THE  WARRANTIES OF  MERCHANTABILITY,
FITNESS FOR  A PARTICULAR PURPOSE AND  NONINFRINGEMENT. IN NO EVENT  SHALL THE
AUTHORS  OR COPYRIGHT  HOLDERS  BE  LIABLE FOR  ANY  CLAIM,  DAMAGES OR  OTHER
LIABILITY, WHETHER IN AN ACTION OF  CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE  OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

/**
 * Compile with GCC/MinGW:
 *   x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -D_UNICODE -municode -o shortcutinfo.exe  shortcutinfo.cpp  -lole32 -luuid -static -s
 *
 * Compile with MSVC:
 *   cl.exe -W3 -O2 -D_UNICODE shortcutinfo.cpp
 */

#ifdef _MSC_VER
#pragma comment(lib, "ole32.lib")
#endif
#include <windows.h>
#include <objbase.h>
#include <shlobj.h>
#include <stdio.h>
#include <tchar.h>

#define CAST_VPP(x) reinterpret_cast<void **>(x)


inline void print_guid(unsigned char *p)
{
	_tprintf_s(_T("::{%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X}"),
		p[3], p[2], p[1], p[0],
		p[5] ,p[4],
		p[7], p[6],
		p[8], p[9],
		p[10], p[11], p[12], p[13], p[14], p[15]);
}

#ifdef _UNICODE
#define char_to_wchar(x) x
#else
wchar_t *char_to_wchar(const char *str)
{
	wchar_t *pwcs = NULL;
	size_t len = mbstowcs(NULL, str, 0);

	if (len == static_cast<size_t>(-1)) {
		return NULL;
	}

	pwcs = reinterpret_cast<wchar_t *>(malloc(len + 2));
	mbstowcs(pwcs, str, len + 1);

	return pwcs;
}
#endif

int _tmain(int argc, _TCHAR *argv[])
{
	IShellLink *shLink = NULL;
	IShellLinkDataList *shLdl = NULL;
	IPersistFile *pFile = NULL;
	FILE *fp = NULL;
	DWORD dwFlags = 0;
	int ret = 1;

	const int bufLen = 2048;
	_TCHAR buf[bufLen] = {0};
	_TCHAR *pbuf = buf;
	WORD wHotkey = 0;
	int n = 0;

	const _TCHAR *isUC = (sizeof(_TCHAR) > 1) ? _T(" (unicode build)") : _T("");

	if (argc < 2) {
		_tprintf_s(_TEXT(
				"Shows information about Shell Links%s\n"
				"usage: %s FILENAME\n"), isUC, argv[0]);
		return 0;
	}

	const _TCHAR *filename = argv[1];
	LPCOLESTR olestr = char_to_wchar(argv[1]);

	// initialize COM library
	const DWORD dwCoFlags =
		COINIT_APARTMENTTHREADED |
		COINIT_DISABLE_OLE1DDE |
		COINIT_SPEED_OVER_MEMORY;

	HRESULT hRes = CoInitializeEx(NULL, dwCoFlags);

	if (hRes != S_OK) {
		_tprintf_s(_T("%s: could not initialize COM library\n"), argv[0]);
		return 1;
	}

	// create instance
	hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
			IID_IShellLink, CAST_VPP(&shLink));

	if (hRes != S_OK) {
		goto MAIN_EXIT;
	}

	// query interface
	hRes = shLink->QueryInterface(IID_IPersistFile, CAST_VPP(&pFile));

	if (hRes != S_OK) {
		goto MAIN_EXIT;
	}

	// open file
	if (pFile->Load(olestr, 0) != S_OK) {
		goto MAIN_EXIT;
	}

	// path
	if (shLink->GetPath(pbuf, bufLen, NULL, 0) == S_OK && _tcslen(buf) > 0) {
		_tprintf_s(_T("Target path: %s\n"), buf);
	}

	// CLSID (not the best solution)
	if (_tfopen_s(&fp, filename, _T("rb")) == 0 && fp != NULL &&
		fseek(fp, 76, SEEK_SET) == 0)
	{
		unsigned char data[64] = {0};
		int len = fgetc(fp);

		if (len == 64) {
			// 2 CLSIDs, i.e. ::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{7B81BE6A-CE2B-4676-A29E-EB907A5126C5}
			if (fread(&data, 1, 63, fp) == 63) {
				_tprintf_s(_T("CLSID: "));
				print_guid(data+5);
				_tprintf_s(_T("\\0\\"));
				print_guid(data+47);
				putwchar(_T('\n'));
			}
		} else {
			// 1 CLSID
			if (fread(&data, 1, 21, fp) == 21) {
				_tprintf_s(_T("CLSID: "));
				print_guid(data+5);
				putwchar(_T('\n'));
			}
		}
	}

	// arguments
	if (shLink->GetArguments(pbuf, bufLen) == S_OK && _tcslen(buf) > 0) {
		_tprintf_s(_T("Arguments: %s\n"), buf);
	}

	// description
	if (shLink->GetDescription(pbuf, bufLen) == S_OK && _tcslen(buf) > 0) {
		_tprintf_s(_T("Description: %s\n"), buf);
	}

	// icon
	if (shLink->GetIconLocation(pbuf, bufLen, &n) == S_OK && _tcslen(buf) > 0) {
		_tprintf_s(_TEXT(
				"Icon location: %s\n"
				"Icon index: %d\n"), buf, n);
	}

	// working directory
	if (shLink->GetWorkingDirectory(pbuf, bufLen) == S_OK && _tcslen(buf) > 0) {
		_tprintf_s(_T("Working directory: %s\n"), buf);
	}

	// showcmd
	if (shLink->GetShowCmd(&n) == S_OK) {
		_tprintf_s(_T("Show command: "));

		switch(n) {
			case SW_SHOWNORMAL:
				_tprintf_s(_T("SW_SHOWNORMAL\n"));
				break;
			case SW_SHOWMAXIMIZED:
				_tprintf_s(_T("SW_SHOWMAXIMIZED\n"));
				break;
			case SW_SHOWMINIMIZED:
				_tprintf_s(_T("SW_SHOWMINIMIZED\n"));
				break;
			default:
				_tprintf_s(_T("0x%x (not supported)\n"), n);
				break;
		}
	}

	// hotkey
	if (shLink->GetHotkey(&wHotkey) == S_OK && wHotkey != 0) {
		unsigned char hi = (wHotkey >> 8) & 0xff;
		unsigned char lo = wHotkey & 0xff;

		_tprintf_s(_T("Hotkey:"));

		if (hi & HOTKEYF_CONTROL) _tprintf_s(_T(" [CTRL]"));
		if (hi & HOTKEYF_SHIFT) _tprintf_s(_T(" [SHIFT]"));
		if (hi & HOTKEYF_ALT) _tprintf_s(_T(" [ALT]"));
		if (hi & HOTKEYF_EXT) _tprintf_s(_T(" [EXT]"));

		if ((lo >= 'A' && lo <= 'Z') || (lo >= '0' && lo <= '9')) {
			_tprintf_s(_T(" [%c]"), lo);
		} else if (lo >= VK_F1 && lo <= VK_F24) {
			_tprintf_s(_T(" [F%d]"), lo - (VK_F1 - 1));
		} else if (lo == VK_NUMLOCK) {
			_tprintf_s(_T(" [NUMLOCK]"));
		} else if (lo == VK_SCROLL) {
			_tprintf_s(_T(" [SCROLL]"));
		} else if (lo == 0x00) {
			_tprintf_s(_T(" [0x00 (not set)]"));
		} else {
			_tprintf_s(_T(" [0x%02X (not supported)]"), lo);
		}

		_tprintf_s(_T(" (0x%X)\n"), wHotkey);
	}

	// ShellLinkDataList (for flags)
	hRes = shLink->QueryInterface(IID_IShellLinkDataList, CAST_VPP(&shLdl));

	if (hRes == S_OK && shLdl->GetFlags(&dwFlags) == S_OK) {
		const _TCHAR *yesno = (dwFlags & SLDF_RUNAS_USER) ? _T("yes") : _T("no");
		_tprintf_s(_T("Run as Administrator: %s\n"), yesno);
	}

	ret = 0;

MAIN_EXIT:

#ifndef _UNICODE
	if (olestr) free((void *)olestr);
#endif
	if (fp) fclose(fp);

	if (shLdl) shLdl->Release();
	if (pFile) pFile->Release();
	if (shLink) shLink->Release();
	CoUninitialize();

	return ret;
}
