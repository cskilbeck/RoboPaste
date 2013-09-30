/****************************** Module Header ******************************\
Module Name:  FileContextMenuExt.cpp
Project:      CppShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The code sample demonstrates creating a Shell context menu handler with C++. 

A context menu handler is a shell extension handler that adds commands to an 
existing context menu. Context menu handlers are associated with a particular 
file class and are called any time a context menu is displayed for a member 
of the class. While you can add items to a file class context menu with the 
registry, the items will be the same for all members of the class. By 
implementing and registering such a handler, you can dynamically add items to 
an object's context menu, customized for the particular object.

The example context menu handler adds the menu item "Display File Name (C++)"
to the context menu when you right-click a .cpp file in the Windows Explorer. 
Clicking the menu item brings up a message box that displays the full path 
of the .cpp file.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <Mmsystem.h>
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "winmm.lib")

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

FileContextMenuExt::FileContextMenuExt(void) : m_cRef(1), 
    m_pszMenuText(L"&RoboPaste"),
    m_pszVerb("robopaste"),
    m_pwszVerb(L"robopaste"),
    m_pszVerbCanonicalName("RoboPaste"),
    m_pwszVerbCanonicalName(L"RoboPaste"),
    m_pszVerbHelpText("RoboPaste"),
    m_pwszVerbHelpText(L"RoboPaste")
{
    InterlockedIncrement(&g_cDllRef);

    // Load the bitmap for the menu item. 
    // If you want the menu item bitmap to be transparent, the color depth of 
    // the bitmap must not be greater than 8bpp.
    m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK), IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);
}

FileContextMenuExt::~FileContextMenuExt(void)
{
    if (m_hMenuBmp)
    {
        DeleteObject(m_hMenuBmp);
        m_hMenuBmp = NULL;
    }

    InterlockedDecrement(&g_cDllRef);
}

HHOOK hhk;

std::wstring replacers[3] =
{
	std::wstring(L"Yes!!"),
	std::wstring(L"No!!"),
	std::wstring(L"Cancel!!")
};

LRESULT CALLBACK CBTProc(INT nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HCBT_ACTIVATE)
	{
		HWND h = (HWND)wParam;
		SetDlgItemText(h, IDYES, replacers[0].c_str());
		SetDlgItemText(h, IDNO, replacers[1].c_str());
		SetDlgItemText(h, IDCANCEL, replacers[2].c_str());
		UnhookWindowsHookEx(hhk);
	}
	else
	{
		CallNextHookEx(hhk, nCode, wParam, lParam);
	}
	return 0;
}

int MsgBox(HWND hwnd, WCHAR const *msg, WCHAR const *yes, WCHAR const *no, WCHAR const *cancel = NULL)
{
	replacers[0] = yes;
	replacers[1] = no;
	replacers[2] = (cancel == NULL) ? L"" : cancel;
	hhk = SetWindowsHookEx(WH_CBT, &CBTProc, 0, GetCurrentThreadId());
	return MessageBox(hwnd, msg, L"RoboPaste", (cancel == NULL) ? MB_YESNO : MB_YESNOCANCEL);
}

std::wstring Format(WCHAR const *fmt, ...)
{
	WCHAR buf[1024];
	va_list v;
	va_start(v, fmt);
	_vsnwprintf_s(buf, ARRAYSIZE(buf), fmt, v);
	return std::wstring(buf);
}

bool FileExists(TCHAR const * file)
{
	WIN32_FIND_DATA FindFileData;
	HANDLE handle = FindFirstFile(file, &FindFileData) ;
	bool found = handle != INVALID_HANDLE_VALUE;
	if(found) 
	{
		FindClose(handle);
	}
	return found;
}

void FileContextMenuExt::OnVerbDisplayFileName(HWND hWnd)
{
	WCHAR *personalFolder;
	if(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &personalFolder) == S_OK)
	{
		std::wstring s(personalFolder);
		s += L"\\RoboBatch";
		CoTaskMemFree(personalFolder);
		if(CreateDirectory(s.c_str(), NULL) != 0 || GetLastError() == ERROR_ALREADY_EXISTS)
		{
			Log(L"Temp filename is ");
			bool found = false;
			std::wstring tfn;
			int tries = 0;
			HANDLE file;
			bool fatal = false;
			while(!found && ++tries < 100 && !fatal)
			{
				tfn = Format(L"%s\\robo%08x.bat", s.c_str(), timeGetTime());
				file = CreateFile(tfn.c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
				if(file != INVALID_HANDLE_VALUE)
				{
					found = true;
				}
				else
				{
					if(GetLastError() == ERROR_ALREADY_EXISTS)
					{
						Sleep(1);
					}
					else
					{
						Log(L"Error creating file %s: %08x\n", tfn.c_str(), GetLastError());
						break;
					}
				}
			}
			if(found)
			{
				Log(L"%s\n", tfn.c_str());
				bool error = false;
				for(int fn=0; fn<mFiles.size() && !error; ++fn)
				{
					std::wstring s = mFiles[fn];
					DWORD wrote;
					DWORD fileAttributes = GetFileAttributes(s.c_str());
					if(fileAttributes != INVALID_FILE_ATTRIBUTES)
					{
						std::wstring batchLine;

						WCHAR drive[_MAX_DRIVE];
						WCHAR path[_MAX_DIR];
						WCHAR filename[_MAX_FNAME];
						WCHAR ext[_MAX_EXT];
						
						_wsplitpath_s(s.c_str(), drive, path, filename, ext);

						std::wstring driveStr(drive);
						std::wstring pathStr(path);
						if(pathStr[pathStr.size() - 1] == L'\\')
						{
							pathStr.pop_back();
						}
						std::wstring fileStr(filename);
						std::wstring extStr(ext);

						std::wstring fullPath = driveStr + pathStr;
						std::wstring fullName = fileStr + extStr;

						Log(L"FullPath: %s\nFullName: %s\n", fullPath.c_str(), fullName.c_str());

						bool isCandidate = (fileAttributes & (FILE_ATTRIBUTE_DEVICE | FILE_ATTRIBUTE_OFFLINE)) == 0;
						bool isFolder = (fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && isCandidate;
						if(isFolder)
						{
							batchLine += Format(L"mkdir \"%s\\%s\"\r\n", mDestinationPath.c_str(), filename);
							batchLine += Format(L"robocopy \"%s\" \"%s\\%s\" /E\r\n\r\n", s.c_str(), mDestinationPath.c_str(), filename);
						}
						else if(isCandidate)
						{
							batchLine += Format(L"robocopy \"%s\" \"%s\" \"%s\"\r\n\r\n", fullPath.c_str(), mDestinationPath.c_str(), fullName.c_str());
						}
						else
						{
							batchLine += Format(L"REM skipped \"%s\"\r\n\r\n", s.c_str());
						}

						if(!batchLine.empty())
						{
							Log(L"Batchline: %s\n", batchLine.c_str());
							int nch = WideCharToMultiByte(CP_ACP, 0, batchLine.c_str(), (int)batchLine.size(), NULL, 0, NULL, NULL);
							Log(L"%d\n", nch);
							std::vector<char> mbc;
							mbc.resize(nch);
							WideCharToMultiByte(CP_ACP, 0, batchLine.c_str(), (int)batchLine.size(), &mbc[0], (int)mbc.size(), NULL, NULL);

							DWORD wrote = 0;
							if(!WriteFile(file, &mbc[0], (DWORD)mbc.size(), &wrote, NULL))
							{
								error = true;
								Log(L"Error %08x, wrote %d of %d\n", GetLastError(), wrote, mbc.size());
								MessageBox(hWnd, L"Error writing to Robo batch file", L"RoboPaste", MB_ICONEXCLAMATION);
								break;
							}
						}
					}
					else
					{
						switch (MessageBox(hWnd, Format(L"Error getting file information for %s", s.c_str()).c_str(), L"RoboPaste", MB_ABORTRETRYIGNORE))
						{
						case IDABORT:
							error = true;
							break;
							
						case IDRETRY:
							--fn;
							break;

						case IDIGNORE:
							break;
						};
					}
				}
				CloseHandle(file);

				if(!error)
				{
					switch(MsgBox(hWnd, L"Created robo batch file! What do you want to do?", L"Run it", L"Cancel", L"Show it to me"))
					{
					case IDYES:
						{
							std::wstring params = Format(L"/Q /K %s", tfn.c_str());
							SHELLEXECUTEINFO rSEI = { 0 };
							rSEI.cbSize = sizeof( rSEI );
							rSEI.lpVerb = L"open";
							rSEI.lpFile = L"cmd.Exe";
							rSEI.lpParameters = params.c_str();
							rSEI.lpDirectory = mDestinationPath.c_str();
							rSEI.nShow = SW_NORMAL;
							rSEI.fMask = SEE_MASK_NOCLOSEPROCESS;

							if(!ShellExecuteEx(&rSEI))
							{
								Log(L"Error creating process: GLE = %08x\n", GetLastError());
								MessageBox(hWnd, L"Error executing batch file!", L"Robopaste", MB_ICONEXCLAMATION);
							}
						}
						break;

					case IDNO:
						break;

					case IDCANCEL:
						{
							HINSTANCE rc = ShellExecute(hWnd, L"edit", tfn.c_str(), NULL, NULL, SW_SHOW);
							if((DWORD)rc <= 32)
							{
								Log(L"Error creating process: %08x, GLE = %08x\n", rc, GetLastError());
							}
						}
						break;
					}
				}
			}
			else
			{
				Log(L"Can't create file!\n");
				MessageBox(hWnd, L"Error creating Robo batch file", L"RoboPaste", MB_ICONEXCLAMATION);
			}
		}
		else
		{
			Log(L"CreateDirectory failed (%s): GLE: %08x\n", s.c_str(), GetLastError());
			MessageBox(hWnd, L"Error creating RoboBatch folder (%s): %08x", L"RoboPaste", MB_ICONEXCLAMATION);
		}
	}
	else
	{
		MessageBox(hWnd, L"Error finding AppData folder", L"RoboPaste", MB_ICONEXCLAMATION);
	}

}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP FileContextMenuExt::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(FileContextMenuExt, IContextMenu),
        QITABENT(FileContextMenuExt, IShellExtInit), 
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IShellExtInit

// Initialize the context menu handler.
IFACEMETHODIMP FileContextMenuExt::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
	if(pidlFolder == NULL)
	{
		return ERROR_BAD_PATHNAME;
	}
	WCHAR path[MAX_PATH * 5];
	if(SHGetPathFromIDList(pidlFolder, path))
	{
		mDestinationPath = path;
		return S_OK;
	}
	return ERROR_BAD_PATHNAME;
}

#pragma endregion


#pragma region IContextMenu

//
//   FUNCTION: FileContextMenuExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP FileContextMenuExt::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

	// Check the clipboard
	// If it's got some filenames in it
	// Enable the RoboPaste

	mFiles.clear();

	bool gotIt = false;

	LPDATAOBJECT pDO;
	
	if(OleGetClipboard(&pDO) == S_OK)
	{
		FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM stgm;
		if(pDO->GetData(&fmt, &stgm) == S_OK)
		{
			TCHAR path[MAX_PATH * 2];
			UINT nFiles = DragQueryFile((HDROP)stgm.hGlobal, -1, NULL, 0);
			Log(L"%d files\n", nFiles);
			for(UINT i=0; i<nFiles; ++i)
			{
				DragQueryFile((HDROP)stgm.hGlobal, i, path, ARRAYSIZE(path));
				Log(L"%s\n", path);
				mFiles.push_back(path);
			}
		}
		pDO->Release();
	}

	// Use either InsertMenu or InsertMenuItem to add menu items.
    // Learn how to add sub-menu from:
    // http://www.codeproject.com/KB/shell/ctxextsubmenu.aspx

    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    mii.dwTypeData = m_pszMenuText;
    mii.fState = mFiles.empty() ? MFS_DISABLED : MFS_ENABLED;
    mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);
    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &sep))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
    // Set the code value to the offset of the largest command identifier 
    // that was assigned, plus one (1).
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 1));
}


//
//   FUNCTION: FileContextMenuExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP FileContextMenuExt::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
    // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
    // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
    // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
    // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
    // and has additional members that allow Unicode strings to be passed.
    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    // Determines whether the command is identified by its offset or verb.
    // There are two ways to identify commands:
    // 
    //   1) The command's verb string 
    //   2) The command's identifier offset
    // 
    // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
    // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
    // holds a verb string. If the high-order word is zero, the command 
    // offset is in the low-order word of lpcmi->lpVerb.

    // For the ANSI case, if the high-order word is not zero, the command's 
    // verb string is in lpcmi->lpVerb. 
    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // For the Unicode case, if the high-order word is not zero, the 
    // command's verb string is in lpcmi->lpVerbW. 
    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // If the command cannot be identified through the verb string, then 
    // check the identifier offset.
    else
    {
        // Is the command identifier offset supported by this context menu 
        // extension?
        if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP FileContextMenuExt::GetCommandString(UINT_PTR idCommand, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

#pragma endregion