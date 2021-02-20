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
	IShellLinkDataList *shLdl = NULL;
	IPersistFile *pFile = NULL;
	HRESULT hRes;
	DWORD dwFlags = 0;
	int ret = 1;
	wchar_t *filename;

	wchar_t buf[4096] = {0};
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

	filename = argv[1];

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
	if (pFile->Load(filename, 0) != S_OK) {
		goto MAIN_EXIT;
	}

	// path
	if (shLink->GetPath(pbuf, bufLen, NULL, 0) == S_OK) {
		if (wcslen(buf) > 0) {
			wprintf_s(L"Target path: %s\n", buf);
		}
	} else {
		// CLSID (should find a better way to do this)
		FILE *fp = NULL;
		unsigned char data[98] = {0};
		unsigned char *p = data;
		const unsigned char headerData[] = {
			0x4C,0x00,0x00,0x00,0x01,0x14,0x02,0x00,0x00,0x00,0x00,0x00,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46
		};

		if (_wfopen_s(&fp, filename, L"rb") == 0 && fp
			&& fread(&data, 1, sizeof(data), fp) == sizeof(data)
			&& memcmp(data, headerData, sizeof(headerData)) == 0)
		{
			p += sizeof(data) - 16;
			wprintf_s(L"Target path: ::{");
			wprintf_s(L"%02X%02X%02X%02X-", p[3], p[2], p[1], p[0]);
			wprintf_s(L"%02X%02X-", p[5], p[4]);
			wprintf_s(L"%02X%02X-", p[7], p[6]);
			wprintf_s(L"%02X%02X-", p[8], p[9]);
			wprintf_s(L"%02X%02X%02X%02X%02X%02X}\n", p[10], p[11], p[12], p[13], p[14], p[15]);
		}

		if (fp) {
			fclose(fp);
		}
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
			wprintf_s(L" [CTRL]");
		}

		if (hi & HOTKEYF_SHIFT) {
			wprintf_s(L" [SHIFT]");
		}

		if (hi & HOTKEYF_ALT) {
			wprintf_s(L" [ALT]");
		}

		if (hi & HOTKEYF_EXT) {
			wprintf_s(L" [EXT]");
		}

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

	// ShellLinkDataList
	hRes = shLink->QueryInterface(IID_IShellLinkDataList, reinterpret_cast<void **>(&shLdl));

	if (hRes == S_OK && shLdl->GetFlags(&dwFlags) == S_OK) {
		wprintf_s(L"Run as Administrator: ");
		wprintf_s((dwFlags & SLDF_RUNAS_USER) ? L"yes\n" : L"no\n");

/*	EXP_DARWIN_LINK *darw = NULL;
		if ((dwFlags & SLDF_HAS_DARWINID) && shLdl->CopyDataBlock(EXP_DARWIN_ID_SIG, reinterpret_cast<void **>(darw)) == S_OK) {
			wprintf_s(L"szwDarwinID: %s\n", darw->szwDarwinID);
			LocalFree(darw);
		}*/
	}

	ret = 0;

MAIN_EXIT:

	if (shLdl) {
		shLdl->Release();
	}

	if (shLink) {
		shLink->Release();
	}

	if (pFile) {
		pFile->Release();
	}

	CoUninitialize();

	return ret;
}
