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

#pragma once

#include <windows.h>
#include <objbase.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>


class shell_link
{
private:

	const wchar_t *m_filename = NULL;    // Path to shell link (shortcut)
	const wchar_t *m_linktarget = NULL;  // Path to shortcut target (can be a CLSID)
	const wchar_t *m_args = NULL;        // Command line arguments to use on launch
	const wchar_t *m_iconpath = NULL;    // Path to file containing icon (.ico, .icl, .exe, .dll)
	int m_iconidx = 0;                   // Icon index number
	const wchar_t *m_desc = NULL;        // Description
	const wchar_t *m_wdir = NULL;        // Working directory to run command
	int m_showcmd = SW_SHOWNORMAL;       // Show window setting: SW_SHOWNORMAL, SW_SHOWMAXIMIZED or SW_SHOWMINNOACTIVE

	WORD m_hotkey = 0;  // Sets a keyboard shortcut (hot key);
	                    // HighByte modifier flags: HOTKEYF_ALT, HOTKEYF_CONTROL, HOTKEYF_EXT, HOTKEYF_SHIFT
	                    // LowByte key values: 0-9, A-Z, VK_F1-VK_F24, VK_NUMLOCK, VK_SCROLL
	                    // see part 2.1.3 of https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-SHLLINK/%5bMS-SHLLINK%5d.pdf

	bool m_admin = false;                // Flag shell link to be run as Administrator
	HRESULT m_cominitialized = -1;       // Whether COM was initialized or not

	IShellLink *m_shlink = NULL;
	IShellLinkDataList *m_shldl = NULL;
	IPersistFile *m_pfile = NULL;


public:

	shell_link()
	{}

	~shell_link() {
		clear();
	}

	void clear()
	{
		if (m_shldl) m_shldl->Release();
		if (m_pfile) m_pfile->Release();
		if (m_shlink) m_shlink->Release();

		if (SUCCEEDED(m_cominitialized)) {
			CoUninitialize();
		}

		m_shldl = NULL;
		m_pfile = NULL;
		m_shlink = NULL;
		m_cominitialized = -1;
	}

	void filename(const wchar_t *path) { m_filename = path; }
	void linktarget(const wchar_t *path) { m_linktarget = path; }
	void args(const wchar_t *str) { m_args = str; }
	void iconpath(const wchar_t *path) { m_iconpath = path; }
	void description(const wchar_t *str) { m_desc = str; }
	void workingdir(const wchar_t *path) { m_wdir = path; }
	void admin(bool b) { m_admin = b; }

	bool iconidx(const wchar_t *p)
	{
		int n = 0;

		if (p && swscanf_s(p, L"%d", &n) == 1) {
			m_iconidx = n;
			return true;
		}

		return false;
	}

	void showcmd(int sw)
	{
		switch (sw)
		{
		case SW_SHOWMAXIMIZED:
		case SW_SHOWMINNOACTIVE:
			m_showcmd = sw;
			break;
		default:
			m_showcmd = SW_SHOWNORMAL;
			break;
		}
	}

	bool hotkey(const wchar_t *p)
	{
		WORD combo;
		int f = 0;

		if (!p || wcslen(p) < 3) {
			return false;
		}

		if (_wcsnicmp(p, L"ca", 2) == 0) {
			combo = HOTKEYF_CONTROL | HOTKEYF_ALT;
		} else if (_wcsnicmp(p, L"cs", 2) == 0) {
			combo = HOTKEYF_CONTROL | HOTKEYF_SHIFT;
		} else if (_wcsnicmp(p, L"sa", 2) == 0) {
			combo = HOTKEYF_SHIFT | HOTKEYF_ALT;
		} else {
			return false;
		}

		p += 2;

		wchar_t key = towupper(p[0]);

		if (wcslen(p) == 1) {
			// A-Z, 0-9
			if (key >= L'A' && key <= L'Z') {
				m_hotkey = (combo << 8) | ( L'A' + (key - L'A') );
				return true;
			} else if (key >= L'0' && key <= L'9') {
				m_hotkey = (combo << 8) | ( L'0' + (key - L'0') );
				return true;
			}
		} else {
			// Numlock, Scroll, F-keys
			if (_wcsicmp(p, L"numlock") == 0) {
				m_hotkey = (combo << 8) | VK_NUMLOCK;
				return true;
			} else if (_wcsicmp(p, L"scroll") == 0) {
				m_hotkey = (combo << 8) | VK_SCROLL;
				return true;
			} else if (key == L'F' && swscanf_s(p+1, L"%d", &f) == 1 &&
						f >= 1 && f <= 24)
			{
				m_hotkey = (combo << 8) | (VK_F1 + (f - 1));
				return true;
			}
		}

		return false;
	}

	bool create()
	{
		clear();

		// filename and link target required
		if (!m_filename || !m_linktarget) {
			return false;
		}

		// initialize COM library
		const DWORD dwCoFlags =
			COINIT_APARTMENTTHREADED |
			COINIT_DISABLE_OLE1DDE |
			COINIT_SPEED_OVER_MEMORY;

		m_cominitialized = CoInitializeEx(NULL, dwCoFlags);

		if (FAILED(m_cominitialized)) {
			return false;
		}

		// create instance
		if (FAILED(CoCreateInstance(CLSID_ShellLink,
									NULL,
									CLSCTX_INPROC_SERVER,
									IID_IShellLink,
									reinterpret_cast<void **>(&m_shlink))))
		{
			return false;
		}

		// query interface
		if (FAILED(m_shlink->QueryInterface(IID_IPersistFile,
											reinterpret_cast<void **>(&m_pfile))))
		{
			return false;
		}

		// create Shell Link file
		if (                  FAILED(m_shlink->SetPath(m_linktarget)) ||
			(m_args        && FAILED(m_shlink->SetArguments(m_args))) ||
			(m_iconpath    && FAILED(m_shlink->SetIconLocation(m_iconpath, m_iconidx))) ||
			(m_desc        && FAILED(m_shlink->SetDescription(m_desc))) ||
			(m_wdir        && FAILED(m_shlink->SetWorkingDirectory(m_wdir))) ||
			                  FAILED(m_shlink->SetShowCmd(m_showcmd)) ||
			(m_hotkey != 0 && FAILED(m_shlink->SetHotkey(m_hotkey))))
		{
			return false;
		}

		// set SLDF_RUNAS_USER flag
		if (m_admin) {
			DWORD dwFlags = 0;

			if (FAILED(m_shlink->QueryInterface(IID_IShellLinkDataList,
												reinterpret_cast<void **>(&m_shldl))) ||
				FAILED(m_shldl->GetFlags(&dwFlags)) ||
				FAILED(m_shldl->SetFlags(SLDF_RUNAS_USER | dwFlags)))
			{
				return false;
			}
		}

		// save Shell Link file
		if (SUCCEEDED(m_pfile->Save(m_filename, TRUE)) &&
			SUCCEEDED(m_pfile->SaveCompleted(m_filename)))
		{
			return true;
		}

		return false;
	}
};
