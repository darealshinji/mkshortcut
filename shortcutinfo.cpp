/*
Licensed under the MIT License <http://opensource.org/licenses/MIT>.
SPDX-License-Identifier: MIT
Copyright (c) 2020-2026 djcj@gmx.de

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
#include <stdio.h>
#include "shortcutinfo.hpp"


int wmain(int argc, wchar_t *argv[])
{
	WORD wHotkey = 0;
	DWORD dwFlags = 0;
	const wchar_t *p = NULL;
	int n = 0;

	if (argc < 2) {
		wprintf_s(L"Shows information about Shell Links\n"
					"usage: %s FILENAME\n", argv[0]);
		return 0;
	}

	const wchar_t *filename = argv[1];
	shell_link_info shl(filename);

	if (!shl.load_file()) {
		wprintf_s(L"%s: failed to load file: %s\n", argv[0], filename);
		return 1;
	}

	if ((p = shl.get_path()) != NULL) {
		wprintf_s(L"Target path: %s\n", p);
	}

	if ((p = shl.get_clsid()) != NULL) {
		wprintf_s(L"CLSID: %s\n", p);
	}

	if ((p = shl.get_arguments()) != NULL) {
		wprintf_s(L"Arguments: %s\n", p);
	}

	if ((p = shl.get_description()) != NULL) {
		wprintf_s(L"Description: %s\n", p);
	}

	if ((p = shl.get_iconlocation(n)) != NULL) {
		wprintf_s(L"Icon location: %s\nIcon index: %d\n", p, n);
	}

	if ((p = shl.get_workingdir()) != NULL) {
		wprintf_s(L"Working directory: %s\n", p);
	}

	if (shl.get_showcmd(n)) {
		wprintf_s(L"Show command: ");

		switch(n)
		{
		case SW_SHOWNORMAL:
			wprintf_s(L"SW_SHOWNORMAL\n");
			break;
		case SW_SHOWMAXIMIZED:
			wprintf_s(L"SW_SHOWMAXIMIZED\n");
			break;
		case SW_SHOWMINNOACTIVE:
			wprintf_s(L"SW_SHOWMINNOACTIVE\n");
			break;
		default:
			wprintf_s(L"0x%x (not supported)\n", n);
			break;
		}
	}

	if (shl.get_hotkey(wHotkey) && wHotkey != 0) {
		unsigned char hi = (wHotkey >> 8) & 0xff;
		unsigned char lo = wHotkey & 0xff;

		wprintf_s(L"Hotkey:");

		if (hi & HOTKEYF_CONTROL) wprintf_s(L" [CTRL]");
		if (hi & HOTKEYF_SHIFT) wprintf_s(L" [SHIFT]");
		if (hi & HOTKEYF_ALT) wprintf_s(L" [ALT]");
		if (hi & HOTKEYF_EXT) wprintf_s(L" [EXT]");

		if ((lo >= 'A' && lo <= 'Z') || (lo >= '0' && lo <= '9')) {
			wprintf_s(L" [%c]", lo);
		} else if (lo >= VK_F1 && lo <= VK_F24) {
			wprintf_s(L" [F%d]", lo - (VK_F1 - 1));
		} else if (lo == VK_NUMLOCK) {
			wprintf_s(L" [NUMLOCK]");
		} else if (lo == VK_SCROLL) {
			wprintf_s(L" [SCROLL]");
		} else if (lo == 0x00) {
			wprintf_s(L" [0x00 (not set)]");
		} else {
			wprintf_s(L" [0x%02X (not supported)]", lo);
		}

		wprintf_s(L" (0x%X)\n", wHotkey);
	}

	if (shl.get_flags(dwFlags)) {
		wprintf_s(L"Run as Administrator: %s\n",
					(dwFlags & SLDF_RUNAS_USER) ? L"yes" : L"no");
	}

	return 0;
}
