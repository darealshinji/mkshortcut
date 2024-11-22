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
 * Create a Shell Link a.k.a. Shortcut from command line
 *
 * Compile with GCC/MinGW:
 *   x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -D_UNICODE -DUNICODE -municode -o mkshortcut.exe  mkshortcut.cpp  -lole32 -luuid -static -s
 *
 * Compile with MSVC:
 *   cl.exe -W3 -O2 -D_UNICODE -DUNICODE mkshortcut.cpp
 */

#ifdef _MSC_VER
# define _CRT_SECURE_NO_WARNINGS
# pragma comment(lib, "ole32.lib")
#endif
#include <windows.h>
#include <objbase.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>
#include <tchar.h>

#define CAST_VPP(x) reinterpret_cast<void **>(x)


typedef struct {
	const _TCHAR *pszFileName;    // Path to shell link (shortcut)
	const _TCHAR *pszLinkTarget;  // Path to shortcut target (can be a CLSID)
	const _TCHAR *pszArgs;        // Command line arguments to use on launch
	const _TCHAR *pszIconPath;    // Path to file containing icon (.ico, .icl, .exe, .dll)
	int iIcon;                    // Icon index number
	const _TCHAR *pszDesc;        // Description
	const _TCHAR *pszDir;         // Working directory to run command
	int iShowCmd;                 // Show window setting: SW_SHOWNORMAL, SW_SHOWMAXIMIZED or SW_SHOWMINNOACTIVE
	WORD wHotKey;                 // Sets a keyboard shortcut (hot key);
	                              // HighByte modifier flags: HOTKEYF_ALT, HOTKEYF_CONTROL, HOTKEYF_EXT, HOTKEYF_SHIFT
	                              // LowByte key values: 0-9, A-Z, VK_F1-VK_F24, VK_NUMLOCK, VK_SCROLL
	                              // see part 2.1.3 of https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-SHLLINK/%5bMS-SHLLINK%5d.pdf
	bool bAdmin;                  // Flag shell link to be run as Administrator
} shell_link_t;


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

bool create_shell_link(shell_link_t *opts)
{
	IShellLink *shLink = NULL;
	IShellLinkDataList *shLdl = NULL;
	IPersistFile *pFile = NULL;
	bool ret = false;

	// filename and link target required
	if (!opts->pszFileName || !opts->pszLinkTarget) {
		return false;
	}

	LPCOLESTR olestr = char_to_wchar(opts->pszFileName);

	// initialize COM library
	const DWORD dwCoFlags =
		COINIT_APARTMENTTHREADED |
		COINIT_DISABLE_OLE1DDE |
		COINIT_SPEED_OVER_MEMORY;

	HRESULT hRes = CoInitializeEx(NULL, dwCoFlags);

	if (hRes != S_OK && hRes != S_FALSE) {
		return false;
	}

	// create instance
	hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
			IID_IShellLink, CAST_VPP(&shLink));

	if (hRes != S_OK) {
		CoUninitialize();
		return false;
	}

	// query interface
	hRes = shLink->QueryInterface(IID_IPersistFile, CAST_VPP(&pFile));

	if (hRes != S_OK) {
		shLink->Release();
		CoUninitialize();
		return false;
	}

	// make sure iShowCmd has only one of the three allowed values
	switch(opts->iShowCmd)
	{
		case SW_SHOWNORMAL:
		case SW_SHOWMAXIMIZED:
		case SW_SHOWMINNOACTIVE:
			break;
		default:
			opts->iShowCmd = SW_SHOWNORMAL;
			break;
	}

	// create Shell Link file
	if (                          shLink->SetPath(opts->pszLinkTarget) == S_OK
		&& (!opts->pszArgs     || shLink->SetArguments(opts->pszArgs) == S_OK)
		&& (!opts->pszIconPath || shLink->SetIconLocation(opts->pszIconPath, opts->iIcon) == S_OK)
		&& (!opts->pszDesc     || shLink->SetDescription(opts->pszDesc) == S_OK)
		&& (!opts->pszDir      || shLink->SetWorkingDirectory(opts->pszDir) == S_OK)
		&&                        shLink->SetShowCmd(opts->iShowCmd) == S_OK
		&& (opts->wHotKey == 0 || shLink->SetHotkey(opts->wHotKey) == S_OK))
	{
		ret = true;
	} else {
		pFile->Release();
		shLink->Release();
		CoUninitialize();
		return false;
	}

	// set SLDF_RUNAS_USER flag
	if (opts->bAdmin) {
		DWORD dwFlags = 0;
		hRes = shLink->QueryInterface(IID_IShellLinkDataList, CAST_VPP(&shLdl));

		if (hRes != S_OK || shLdl->GetFlags(&dwFlags) != S_OK
			|| shLdl->SetFlags(SLDF_RUNAS_USER | dwFlags) != S_OK)
		{
			ret = false;
		}
	}

	// save Shell Link file
	if (ret == true && pFile->Save(olestr, TRUE) == S_OK) {
		pFile->SaveCompleted(olestr);
	} else {
		ret = false;
	}

	// free resources
#ifndef _UNICODE
	if (olestr) free((void *)olestr);
#endif
	if (shLdl) shLdl->Release();
	pFile->Release();
	shLink->Release();
	CoUninitialize();

	return ret;
}

bool hotkey_from_string(const _TCHAR *p, WORD &wHotKey)
{
	WORD combo;
	_TCHAR key;
	int f = 0;

	if (_tcsncicmp(p, _T("ca"), 2) == 0) {
		combo = HOTKEYF_CONTROL | HOTKEYF_ALT;
	} else if (_tcsncicmp(p, _T("cs"), 2) == 0) {
		combo = HOTKEYF_CONTROL | HOTKEYF_SHIFT;
	} else if (_tcsncicmp(p, _T("sa"), 2) == 0) {
		combo = HOTKEYF_SHIFT | HOTKEYF_ALT;
	} else {
		return false;
	}

	p += 2;
	key = toupper(p[0]);

	if (_tcslen(p) == 1) {
		// A-Z, 0-9
		if (key >= _T('A') && key <= _T('Z')) {
			wHotKey = (combo << 8) | ( _T('A') + (key - _T('A')) );
			return true;
		} else if (key >= _T('0') && key <= _T('9')) {
			wHotKey = (combo << 8) | ( _T('0') + (key - _T('0')) );
			return true;
		} else {
			return false;
		}
	} else {
		// Numlock, Scroll, F-keys
		if (_tcsicmp(p, _T("numlock")) == 0) {
			wHotKey = (combo << 8) | VK_NUMLOCK;
			return true;
		} else if (_tcsicmp(p, _T("scroll")) == 0) {
			wHotKey = (combo << 8) | VK_SCROLL;
			return true;
		} else if (key == _T('F') && _stscanf_s(p+1, _T("%d"), &f) == 1) {
			if (f < 1 || f > 24) {
				return false;
			}
			wHotKey = (combo << 8) | (VK_F1 + (f - 1));
			return true;
		} else {
			return false;
		}
	}

	return true;
}

static const _TCHAR *help_text = _TEXT(
	"Create a Shell Link a.k.a. Shortcut%s\n"
	"\n"
	"Usage: %s [options]\n"
	"\n"
	"  Options can begin with '/' or '-' and are case-insensitive,\n"
	"  argument separator can be ':' or '='\n"
	"\n"
	"Options:\n"
	"  /? or /h or /help   Print this text\n"
	"\n"
	"  /o:<output>         Path to shell link (shortcut); should end on .lnk [mandatory]\n"
	"  /t:<target>         Path to shortcut target or CLSID [mandatory]\n"
	"                      To set a CLSID the target parameter must be\n"
	"                      set as ::{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\n"
	"  /a:<arguments>      Command line arguments to use on launch\n"
	"  /i:<icon>           Path to file containing icon (.ico, .icl, .exe, .dll)\n"
	"  /n:<index>          Icon index number\n"
	"  /d:<description>    Description (for tooltip)\n"
	"  /w:<directory>      Working directory to run command\n"
	"  /k:<hotkey>         Set hotkey (cak, csk or sak = Ctrl+Alt+Key,\n"
	"                      Ctrl+Shift+Key or Shift+Alt+Key; Key must be 0-9,\n"
	"                      A-Z, F1-F24, NUMLOCK or SCROLL)\n"
	"  /max                Start with maximized window\n"
	"  /min                Start with minimized window\n"
	"  /tfull              Resolve path to shortcut target to a full path\n"
	"  /ifull              Resolve path to icon file to a full path\n"
	"  /admin              Flag shortcut to be run as Administrator\n"
	"\n");


int _tmain(int argc, _TCHAR *argv[])
{
	const _TCHAR *p = NULL;
	const _TCHAR *hkOpt = NULL;
	const _TCHAR *prog = argv[0];
	_TCHAR *fullPathLink = NULL;
	_TCHAR *fullPathTarget = NULL;
	_TCHAR *fullPathIcon = NULL;
	int ret = 0;
	bool tFull = false;
	bool iFull = false;

	const _TCHAR *invOptMsg = _T("%s: invalid option -- '%s'\nTry '%s /?' for more information.\n");
	const _TCHAR *isUC = (sizeof(_TCHAR) > 1) ? _T(" (unicode build)") : _T("");

	if (argc < 2) {
		_tprintf_s(help_text, isUC, prog);
		return 1;
	}

	shell_link_t opts;
	memset(&opts, 0, sizeof(shell_link_t));
	opts.iShowCmd = SW_SHOWNORMAL;

	// parse arguments
	for (int i = 1; i < argc; ++i) {
		const _TCHAR *a = argv[i];

		// all args should start with '/' or '-'
		if ((a[0] != _T('/') && a[0] != _T('-')) || _tcslen(a) < 2) {
			_tprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		// print help
		if (_tcscmp(a+1, _T("?")) == 0 || _tcsicmp(a+1, _T("h")) == 0 || _tcsicmp(a+1, _T("help")) == 0) {
			_tprintf_s(help_text, isUC, prog);
			return 0;
		}

		// from here on length should be at least 4
		if (_tcslen(a) < 4) {
			_tprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		if (_tcsicmp(a+1, _T("max")) == 0) {
			opts.iShowCmd = SW_SHOWMAXIMIZED;
			continue;
		} else if (_tcsicmp(a+1, _T("min")) == 0) {
			opts.iShowCmd = SW_SHOWMINNOACTIVE;
			continue;
		} else if (_tcsicmp(a+1, _T("tfull")) == 0) {
			tFull = true;
			continue;
		} else if (_tcsicmp(a+1, _T("ifull")) == 0) {
			iFull = true;
			continue;
		} else if (_tcsicmp(a+1, _T("admin")) == 0) {
			opts.bAdmin = true;
			continue;
		}

		// from here on argument pattern should be '/x:[...]'
		if (a[2] != _T(':') && a[2] != _T('=')) {
			_tprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		switch(_totlower(a[1])) {
			case _T('o'):
				opts.pszFileName = a+3;
				break;
			case _T('t'):
				opts.pszLinkTarget = a+3;
				break;
			case _T('a'):
				opts.pszArgs = a+3;
				break;
			case _T('i'):
				opts.pszIconPath = a+3;
				break;
			case _T('n'):
				if (_stscanf_s(a+3, _T("%d"), &opts.iIcon) != 1) {
					_tprintf_s(invOptMsg, prog, a, prog);
					return 1;
				}
				break;
			case _T('d'):
				opts.pszDesc = a+3;
				break;
			case _T('w'):
				opts.pszDir = a+3;
				break;
			case _T('k'):
				hkOpt = a;
				break;
			default:
				_tprintf_s(invOptMsg, prog, a, prog);
				return 1;
		}
	}

	// check if pszFileName was set
	if (!opts.pszFileName) {
		_tprintf_s(_T("%s: no output given\nTry '%s /?' for more information.\n"), prog, prog);
		return 1;
	}

	// check if pszLinkTarget was set
	if (!opts.pszLinkTarget) {
		_tprintf_s(_T("%s: no target given\nTry '%s /?' for more information.\n"), prog, prog);
		return 1;
	}

	// hotkey
	if (hkOpt && (_tcslen(hkOpt) < 6 || !hotkey_from_string(hkOpt+3, opts.wHotKey))) {
		_tprintf_s(invOptMsg, prog, hkOpt, prog);
		return 1;
	}

	// warn if output doesn't end on .lnk
	p = _tcsrchr(opts.pszFileName, _T('.'));

	if (!p || _tcsicmp(p, _T(".lnk")) != 0) {
		_tprintf_s(_T("Warning: output link name doesn't end on '.lnk'!\n\n"));
	}

	// make full paths

	if (tFull) {
		fullPathTarget = _tfullpath(NULL, opts.pszLinkTarget, 0);

		if (fullPathTarget) {
			opts.pszLinkTarget = fullPathTarget;
		} else {
			_tprintf_s(_T("%s: failed to resolve full path: %s\n"), prog, opts.pszLinkTarget);
			return 1;
		}
	}

	if (iFull && opts.pszIconPath) {
		fullPathIcon = _tfullpath(NULL, opts.pszIconPath, 0);

		if (fullPathIcon) {
			opts.pszIconPath = fullPathIcon;
		} else {
			_tprintf_s(_T("%s: failed to resolve full path: %s\n"), prog, opts.pszIconPath);
			ret = 1;
			goto MAIN_EXIT;
		}
	}

	// create Shortcut
	if (create_shell_link(&opts)) {
		fullPathLink = _tfullpath(NULL, opts.pszFileName, 0);
		_tprintf_s(_T("Shortcut created:\n%s\n"), fullPathLink ? fullPathLink : opts.pszFileName);
	} else {
		_tprintf_s(_T("%s: create_shell_link() failed\n"), prog);
		ret = 1;
	}

MAIN_EXIT:

	// free resources

	if (fullPathTarget) {
		free(fullPathTarget);
	}

	if (fullPathLink) {
		free(fullPathLink);
	}

	if (fullPathIcon) {
		free(fullPathIcon);
	}

	return ret;
}
