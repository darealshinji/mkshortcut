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
#include <wchar.h>


class shell_link_info
{
private:
	const wchar_t *m_filename = NULL;
	HRESULT m_cominitialized = -1;
	IShellLink *m_shlink = NULL;
	IShellLinkDataList *m_shldl = NULL;
	IPersistFile *m_pfile = NULL;
	wchar_t m_buf[32*1024] = {0};
	wchar_t *m_pbuf = m_buf;


public:

	shell_link_info(const wchar_t *filename)
	: m_filename(filename)
	{}

	~shell_link_info() {
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
		m_buf[0] = 0;
	}

	bool load_file()
	{
		const DWORD dwFlags =
			COINIT_APARTMENTTHREADED |
			COINIT_DISABLE_OLE1DDE |
			COINIT_SPEED_OVER_MEMORY;

		clear();

		m_cominitialized = CoInitializeEx(NULL, dwFlags);

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

		// query interface (persist file)
		if (FAILED(m_shlink->QueryInterface(IID_IPersistFile,
											reinterpret_cast<void **>(&m_pfile))))
		{
			return false;
		}

		// load file
		if (SUCCEEDED(m_pfile->Load(m_filename, 0))) {
			return true;
		}

		return false;
	}

	const wchar_t *get_path()
	{
		if (m_shlink &&
			SUCCEEDED(m_shlink->GetPath(m_pbuf, _countof(m_buf), NULL, 0)) &&
			m_buf[0] != 0)
		{
			return m_pbuf;
		}

		return NULL;
	}

/*
	const wchar_t *get_clsid()
	{
		int n = -1;
		PIDLIST_ABSOLUTE pidl = NULL;

		if (!m_shlink || m_shlink->GetIDList((PIDLIST_ABSOLUTE *)&pidl) != S_OK || !pidl) {
			return NULL;
		}

		USHORT abID_size = pidl->mkid.cb - sizeof(pidl->mkid.cb);

		if (abID_size >= 18) {
			unsigned char *p = pidl->mkid.abID + 2; // 2 bytes for size?

			n = swprintf_s(m_buf, _countof(m_buf) - 1,
							L"::{%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
							p[3], p[2], p[1], p[0],
							p[5] ,p[4],
							p[7], p[6],
							p[8], p[9],
							p[10], p[11], p[12], p[13], p[14], p[15]);
		}

		ILFree(pidl);

		return (n == -1) ? NULL : m_buf;
	}
*/

	const wchar_t *get_arguments()
	{
		if (m_shlink &&
			SUCCEEDED(m_shlink->GetArguments(m_pbuf, _countof(m_buf))) &&
			m_buf[0] != 0)
		{
			return m_pbuf;
		}

		return NULL;
	}

	const wchar_t *get_description()
	{
		if (m_shlink &&
			SUCCEEDED(m_shlink->GetDescription(m_pbuf, _countof(m_buf))) &&
			m_buf[0] != 0)
		{
			return m_pbuf;
		}

		return NULL;
	}

	const wchar_t *get_iconlocation(int &n)
	{
		if (m_shlink &&
			SUCCEEDED(m_shlink->GetIconLocation(m_pbuf, _countof(m_buf), &n)) &&
			m_buf[0] != 0)
		{
			return m_pbuf;
		}

		return NULL;
	}

	const wchar_t *get_workingdir()
	{
		if (m_shlink &&
			SUCCEEDED(m_shlink->GetWorkingDirectory(m_pbuf, _countof(m_buf))) &&
			m_buf[0] != 0)
		{
			return m_pbuf;
		}

		return NULL;
	}

	bool get_showcmd(int &n)
	{
		return (m_shlink && SUCCEEDED(m_shlink->GetShowCmd(&n)));
	}

	bool get_hotkey(WORD &wHotkey)
	{
		return (m_shlink && SUCCEEDED(m_shlink->GetHotkey(&wHotkey)));
	}

	bool get_flags(DWORD &dwFlags)
	{
		if (!m_shlink) {
			return false;
		}

		if (!m_shldl &&
			FAILED(m_shlink->QueryInterface(IID_IShellLinkDataList,
											reinterpret_cast<void **>(&m_shldl))))
		{
			return false;
		}

		return SUCCEEDED(m_shldl->GetFlags(&dwFlags));
	}
};

