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
 * Create a Shell Link a.k.a. Shortcut from command line
 *
 * Compile with GCC/MinGW:
 *   x86_64-w64-mingw32-g++ -Wall -Wextra -O3 -municode -o mkshortcut.exe  mkshortcut.cpp  -lole32 -luuid -static -s
 *
 * Compile with MSVC:
 *   cl.exe /W2 /O2 mkshortcut.cpp
 */

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
#include <stdlib.h>
#include <wchar.h>


bool create_shell_link(
	wchar_t *pszFileName,    // Path to shell link (shortcut)
	wchar_t *pszLinkTarget,  // Path to shortcut target
	wchar_t *pszArgs,        // Command line arguments to use on launch
	wchar_t *pszIconPath,    // Path to file containing icon (.ico, .icl, .exe, .dll)
	int iIcon,               // Icon index number
	wchar_t *pszDesc,        // Description
	wchar_t *pszDir,         // Working directory to run command
	int iShowCmd,            // Show window setting: SW_SHOWNORMAL, SW_SHOWMAXIMIZED or SW_SHOWMINNOACTIVE
	WORD wHotKey             // Sets a keyboard shortcut (hot key); modifier flags are: HOTKEYF_ALT,
	                         // HOTKEYF_CONTROL, HOTKEYF_EXT, HOTKEYF_SHIFT
)
{
	HRESULT hRes;
	IShellLink *shLink;
	IPersistFile *pFile;
	bool ret;

	// filename and link target required
	if (!pszFileName || !pszLinkTarget) {
		return false;
	}

	// initialize OLE Component Object (COM) Library
	hRes = OleInitialize(NULL);

	if (hRes != S_FALSE && hRes != S_OK) {
		return false;
    }

	// create instance
	hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
							IID_IShellLink, reinterpret_cast<void **>(&shLink));

	if (hRes != S_OK) {
		OleUninitialize();
		return false;
	}

	// query interface
	hRes = shLink->QueryInterface(IID_IPersistFile, reinterpret_cast<void **>(&pFile));

	if (hRes != S_OK) {
		shLink->Release();
		OleUninitialize();
		return false;
	}

	// make sure iShowCmd has only one of the three allowed values
	switch(iShowCmd) {
		case SW_SHOWNORMAL:
		case SW_SHOWMAXIMIZED:
		case SW_SHOWMINNOACTIVE:
			break;
		default:
			iShowCmd = SW_SHOWNORMAL;
			break;
	}

	// link target
	if (shLink->SetPath(pszLinkTarget) != S_OK) {
		goto LINK_ERROR;
	}

	// arguments
	if (pszArgs && shLink->SetArguments(pszArgs) != S_OK) {
		goto LINK_ERROR;
	}

	// icon path
	if (pszIconPath && shLink->SetIconLocation(pszIconPath, iIcon) != S_OK) {
		goto LINK_ERROR;
	}

	// description
	if (pszDesc && shLink->SetDescription(pszDesc) != S_OK) {
		goto LINK_ERROR;
	}

	// working directory
	if (pszDir && shLink->SetWorkingDirectory(pszDir) != S_OK) {
		goto LINK_ERROR;
	}

	// show window command
	if (shLink->SetShowCmd(iShowCmd) != S_OK) {
		goto LINK_ERROR;
	}

	// hot key
	if (wHotKey != 0 && shLink->SetHotkey(wHotKey) != S_OK) {
		goto LINK_ERROR;
	}

	// save .lnk file
	if (pFile->Save(pszFileName, FALSE) != S_OK) {
		goto LINK_ERROR;
	}

	// nothing went wrong, return true
	ret = true;
	goto LINK_END;

LINK_ERROR:

	ret = false;

LINK_END:

	// free resources
	pFile->Release();
	shLink->Release();
	OleUninitialize();

	return ret;
}

void print_help(wchar_t *progName)
{
	wprintf_s(
		L"Create a Shell Link a.k.a. Shortcut\n"
		L"\n"
		L"Usage: %s [options]\n"
		L"\n"
		L"  Options can begin with '/' or '-' and are case-insensitive\n"
		L"\n"
		L"Options:\n"
		L"  /? or /h or /help   Print this text\n"
		L"\n"
		L"  /o:<output>         Path to shell link (shortcut); should end on .lnk [mandatory]\n"
		L"  /t:<target>         Path to shortcut target [mandatory]\n"
		L"  /a:<arguments>      Command line arguments to use on launch\n"
		L"  /i:<icon>           Path to file containing icon (.ico, .icl, .exe, .dll)\n"
		L"  /n:<index>          Icon index number\n"
		L"  /d:<description>    Description (for tooltip)\n"
		L"  /w:<directory>      Working directory to run command\n"
		L"  /k:<hotkey>         Set hotkey (ca[a-z], cs[a-z] or sa[a-z] = Ctrl+Alt+[A-Z],\n"
		"                       Ctrl+Shift+[A-Z] or Shift+Alt+[A-Z])\n"
		L"  /max                Start with maximized window\n"
		L"  /min                Start with minimized window\n"
		L"  /tfull              Resolve path to shortcut target to a full path\n"
		L"  /ifull              Resolve path to icon file to a full path\n"
		L"", progName);
}

int wmain(int argc, wchar_t *argv[])
{
	wchar_t *pszFileName = NULL;
	wchar_t *pszLinkTarget = NULL;
	wchar_t *pszArgs = NULL;
	wchar_t *pszIconPath = NULL;
	int iIcon = 0;
	wchar_t *pszDesc = NULL;
	wchar_t *pszDir = NULL;
	int iShowCmd = SW_SHOWNORMAL;
	WORD wHotKey = 0;

	wchar_t *p = NULL;
	wchar_t *hkOpt = NULL;
	wchar_t *prog = argv[0];
	wchar_t *fullPathLink = NULL;
	wchar_t *fullPathTarget = NULL;
	wchar_t *fullPathIcon = NULL;
	int ret = 0;
	bool tFull = false;
	bool iFull = false;

	const wchar_t *invOptMsg = L"%s: invalid option -- '%s'\nTry '%s /?' for more information.\n";


	if (argc < 2) {
		print_help(prog);
		return 1;
	}

	// parse arguments
	for (int i = 1; i < argc; ++i) {
		wchar_t *a = argv[i];

		// all args should start with '/' or '-'
		if ((a[0] != L'/' && a[0] != L'-') || wcslen(a) < 2) {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

        // print help
		if (wcscmp(a+1, L"?") == 0 || _wcsicmp(a+1, L"h") == 0 || _wcsicmp(a+1, L"help") == 0) {
			print_help(prog);
			return 0;
		}

		// from here on length should be at least 4
		if (wcslen(a) < 4) {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		if (_wcsicmp(a+1, L"max") == 0) {
			iShowCmd = SW_SHOWMAXIMIZED;
			continue;
		} else if (_wcsicmp(a+1, L"min") == 0) {
			iShowCmd = SW_SHOWMINNOACTIVE;
			continue;
		} else if (_wcsicmp(a+1, L"tfull") == 0) {
			tFull = true;
			continue;
		} else if (_wcsicmp(a+1, L"ifull") == 0) {
			iFull = true;
			continue;
		}

		// from here on argument pattern should be '/x:[...]'
		if (a[2] != L':') {
			wprintf_s(invOptMsg, prog, a, prog);
			return 1;
		}

		switch(towlower(a[1])) {
			case L'o':
				pszFileName = a+3;
				break;
			case L't':
				pszLinkTarget = a+3;
				break;
			case L'a':
				pszArgs = a+3;
				break;
			case L'i':
				pszIconPath = a+3;
				break;
			case L'n':
				if (swscanf_s(a+3, L"%d", &iIcon) != 1) {
					wprintf_s(invOptMsg, prog, a, prog);
					return 1;
				}
				break;
			case L'd':
				pszDesc = a+3;
				break;
			case L'w':
				pszDir = a+3;
				break;
			case L'k':
				hkOpt = a;
				break;
			default:
				wprintf_s(invOptMsg, prog, a, prog);
				return 1;
		}
	}

	// check if pszFileName was set
	if (!pszFileName) {
		wprintf_s(L"%s: no output given\nTry '%s /?' for more information.\n", prog, prog);
		return 1;
	}

	// check if pszLinkTarget was set
	if (!pszLinkTarget) {
		wprintf_s(L"%s: no target given\nTry '%s /?' for more information.\n", prog, prog);
		return 1;
	}

	// hotkey
	if (hkOpt) {
		WORD combo = 0;
		wchar_t key = 0;

		if (wcslen(hkOpt) != 6) {
			wprintf_s(invOptMsg, prog, hkOpt, prog);
			return 1;
		}

		p = hkOpt + 3;

		if (_wcsnicmp(p, L"ca", 2) == 0) {
			combo = HOTKEYF_CONTROL|HOTKEYF_ALT;
		} else if (_wcsnicmp(p, L"cs", 2) == 0) {
			combo = HOTKEYF_CONTROL|HOTKEYF_SHIFT;
		} else if (_wcsnicmp(p, L"sa", 2) == 0) {
			combo = HOTKEYF_SHIFT|HOTKEYF_ALT;
		}

		key = towlower(p[2]);

		if (combo != 0 && key >= L'a' && key <= L'z') {
			wHotKey = (combo << 8) | ('A' + (key - L'a'));
		} else {
 			wprintf_s(invOptMsg, prog, hkOpt, prog);
			return 1;
		}
	}

	// warn if output doesn't end on .lnk
	p = wcsrchr(pszFileName, L'.');

	if (!p /*|| p == pszFileName*/ || _wcsicmp(p, L".lnk") != 0) {
		wprintf_s(L"Warning: output link name doesn't end on '.lnk'!\n\n");
	}

	// make full paths

	if (tFull) {
		fullPathTarget = _wfullpath(NULL, pszLinkTarget, 0);

		if (fullPathTarget) {
			pszLinkTarget = fullPathTarget;
		} else {
			wprintf_s(L"%s: _wfullpath() failed to resolve path: %s\n", prog, pszLinkTarget);
			return 1;
		}
	}

	if (iFull && pszIconPath) {
		fullPathIcon = _wfullpath(NULL, pszIconPath, 0);

		if (fullPathIcon) {
			pszIconPath = fullPathIcon;
		} else {
			wprintf_s(L"%s: _wfullpath() failed to resolve path: %s\n", prog, pszIconPath);
			ret = 1;
			goto MAIN_EXIT;
		}
	}

	// create Shortcut
	if (create_shell_link(pszFileName, pszLinkTarget, pszArgs, pszIconPath,
							iIcon, pszDesc, pszDir, iShowCmd, wHotKey))
	{
		if ((fullPathLink = _wfullpath(NULL, pszFileName, 0)) != NULL) {
			pszFileName = fullPathLink;
		}
		wprintf_s(L"Shortcut created:\n%s\n", pszFileName);
	} else {
		wprintf_s(L"%s: create_shell_link() failed\n", prog);
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
