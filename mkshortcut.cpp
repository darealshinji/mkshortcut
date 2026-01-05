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
#include "mkshortcut.hpp"


int wmain(int argc, wchar_t *argv[])
{
	const wchar_t *help_text = L""
		"Create a Shell Link a.k.a. Shortcut\n"
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
		"\n";

	const wchar_t *invOptMsg = L""
		"%s: invalid option -- '%s'\n"
		"Try '%s /?' for more information.\n";

	shell_link shlnk;
	const wchar_t *p = NULL;
	const wchar_t *prog = argv[0];
	const wchar_t *pszFileName = NULL;
	const wchar_t *pszLinkTarget = NULL;
	const wchar_t *pszIconPath = NULL;
	wchar_t *fullPathTarget = NULL;
	wchar_t *fullPathIcon = NULL;
	int ret = 0;
	bool tFull = false;
	bool iFull = false;

	if (argc < 2) {
		wprintf_s(help_text, prog);
		return 1;
	}

	// parse arguments
	for (int i = 1; i < argc; ++i) {
		const wchar_t *a = argv[i];

		// all args should start with '/' or '-'
		if ((a[0] != L'/' && a[0] != L'-') || wcslen(a) < 2) {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		// print help
		if (wcscmp(a+1, L"?") == 0 || _wcsicmp(a+1, L"h") == 0 || _wcsicmp(a+1, L"help") == 0) {
			wprintf_s(help_text, prog);
			return 0;
		}

		// from here on length should be at least 4
		if (wcslen(a) < 4) {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		if (_wcsicmp(a+1, L"max") == 0) {
			shlnk.showcmd(SW_SHOWMAXIMIZED);
			continue;
		} else if (_wcsicmp(a+1, L"min") == 0) {
			shlnk.showcmd(SW_SHOWMINNOACTIVE);
			continue;
		} else if (_wcsicmp(a+1, L"tfull") == 0) {
			tFull = true;
			continue;
		} else if (_wcsicmp(a+1, L"ifull") == 0) {
			iFull = true;
			continue;
		} else if (_wcsicmp(a+1, L"admin") == 0) {
			shlnk.admin(true);
			continue;
		}

		// from here on argument pattern should be '/x:[...]'
		if (a[2] != L':' && a[2] != L'=') {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		switch(towlower(a[1])) {
			case L'o':
				pszFileName = a+3;
				shlnk.filename(pszFileName);
				break;
			case L't':
				pszLinkTarget = a+3;
				break;
			case L'a':
				shlnk.args(a+3);
				break;
			case L'i':
				pszIconPath = a+3;
				break;
			case L'n':
				if (!shlnk.iconidx(a+3)) {
					wprintf_s(invOptMsg, prog, a, prog);
					return 1;
				}
				break;
			case L'd':
				shlnk.description(a+3);
				break;
			case L'w':
				shlnk.workingdir(a+3);
				break;
			case L'k':
				if (!shlnk.hotkey(a+3)) {
					wprintf_s(invOptMsg, prog, a, prog);
					return 1;
				}
				break;
			default:
				wprintf_s(invOptMsg, prog, a, prog);
				return 1;
		}
	}

	// check if filename was set
	if (!pszFileName) {
		wprintf_s(L"%s: no output given\n"
					"Try '%s /?' for more information.\n", prog, prog);
		return 1;
	}

	// check if link target was set
	if (!pszLinkTarget) {
		wprintf_s(L"%s: no target given\n"
					"Try '%s /?' for more information.\n", prog, prog);
		return 1;
	}

	// warn if output doesn't end on .lnk
	p = wcsrchr(pszFileName, L'.');

	if (!p || _wcsicmp(p, L".lnk") != 0) {
		wprintf_s(L"Warning: output link name doesn't end on '.lnk'!\n\n");
	}

	// make full paths

	if (tFull) {
		fullPathTarget = _wfullpath(NULL, pszLinkTarget, 0);

		if (fullPathTarget) {
			shlnk.linktarget(fullPathTarget);
		} else {
			wprintf_s(L"%s: failed to resolve full path: %s\n", prog, pszLinkTarget);
			ret = 1;
		}
	}

	if (ret == 0 && iFull && pszIconPath) {
		fullPathIcon = _wfullpath(NULL, pszIconPath, 0);

		if (fullPathIcon) {
			shlnk.iconpath(fullPathIcon);
		} else {
			wprintf_s(L"%s: failed to resolve full path: %s\n", prog, pszIconPath);
			ret = 1;
		}
	}

	// create Shortcut
	if (ret == 0) {
		if (shlnk.create()) {
			wchar_t *buf = _wfullpath(NULL, pszFileName, 0);
			wprintf_s(L"Shortcut created:\n%s\n", buf ? buf : pszFileName);
			free(buf);
		} else {
			wprintf_s(L"%s: failed to create shortcut\n", prog);
			ret = 1;
		}
	}

	free(fullPathIcon);
	free(fullPathTarget);

	return ret;
}
