//////////////////////////////////////////////////////////////////////

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <Mmsystem.h>
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "winmm.lib")

//////////////////////////////////////////////////////////////////////

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

//////////////////////////////////////////////////////////////////////

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
}

//////////////////////////////////////////////////////////////////////

FileContextMenuExt::~FileContextMenuExt(void)
{
    InterlockedDecrement(&g_cDllRef);
}

//////////////////////////////////////////////////////////////////////

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

int CustomMessageBox(HWND hwnd, WCHAR const *msg, WCHAR const *yes, WCHAR const *no, WCHAR const *cancel = NULL)
{
	replacers[0] = yes;
	replacers[1] = no;
	replacers[2] = (cancel == NULL) ? L"" : cancel;
	hhk = SetWindowsHookEx(WH_CBT, &CBTProc, 0, GetCurrentThreadId());
	return MessageBox(hwnd, msg, L"RoboPaste", (cancel == NULL) ? MB_YESNO : MB_YESNOCANCEL);
}

//////////////////////////////////////////////////////////////////////

std::wstring Format(WCHAR const *fmt, ...)
{
	WCHAR buf[1024];
	va_list v;
	va_start(v, fmt);
	_vsnwprintf_s(buf, ARRAYSIZE(buf), fmt, v);
	return std::wstring(buf);
}

//////////////////////////////////////////////////////////////////////

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

//////////////////////////////////////////////////////////////////////

std::vector<WCHAR> GetNonConstString(std::wstring str)
{
	std::vector<WCHAR> v(str.size());
	for(auto c: str)
	{
		v.push_back(c);
	}
	return v;
}

//////////////////////////////////////////////////////////////////////

bool SplitPath(WCHAR const *fullPath, std::wstring &path, std::wstring &fileName)
{
	WCHAR drv[_MAX_DRIVE];
	WCHAR pth[_MAX_DIR];
	WCHAR fnm[_MAX_FNAME];
	WCHAR ext[_MAX_EXT];

	if(_wsplitpath_s(fullPath, drv, pth, fnm, ext) != 0)
	{
		return false;
	}

	std::wstring driveStr(drv);
	std::wstring pathStr(pth);
	if(pathStr[pathStr.size() - 1] == L'\\')
	{
		pathStr.pop_back();
	}
	std::wstring fileStr(fnm);
	std::wstring extStr(ext);

	path = driveStr + pathStr;
	fileName = fileStr + extStr;

	return true;
}

//////////////////////////////////////////////////////////////////////

void BrowseToFile(LPCTSTR filename)
{
	auto pidl = ILCreateFromPath(filename);
	if(pidl)
	{
		SHOpenFolderAndSelectItems(pidl,0,0,0);
		ILFree(pidl);
	}
}

//////////////////////////////////////////////////////////////////////

std::vector<char> WideStringToAnsi(std::wstring str)
{
	std::vector<char> p;
	int nch = WideCharToMultiByte(CP_ACP, 0, str.c_str(), (int)str.size(), NULL, 0, NULL, NULL);
	if(nch < 32767)	// sanity
	{
		p.resize(nch);
		WideCharToMultiByte(CP_ACP, 0, str.c_str(), (int)str.size(), &p[0], nch, NULL, NULL);
	}
	return p;
}

//////////////////////////////////////////////////////////////////////

int ShowError(HWND hWnd, WCHAR const *format, ...)
{
	va_list v;
	va_start(v, format);
	WCHAR buffer[8192];
	_vsnwprintf_s(buffer, ARRAYSIZE(buffer), format, v);
	return MessageBox(hWnd, buffer, L"RoboPaste", MB_ICONEXCLAMATION);
}

//////////////////////////////////////////////////////////////////////

void FileContextMenuExt::OnRoboPaste(HWND hWnd)
{
	// Get %APPDATA%
	WCHAR *personalFolder;
	if(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &personalFolder) == S_OK)
	{
		std::wstring outputPath(personalFolder);
		outputPath += L"\\RoboPaste";
		CoTaskMemFree(personalFolder);

		// Create RoboBatch folder there
		if(CreateDirectory(outputPath.c_str(), NULL) != 0 || GetLastError() == ERROR_ALREADY_EXISTS)
		{
			// Create a new .bat file
			std::wstring batchFilename;
			int tries = 0;
			HANDLE file = INVALID_HANDLE_VALUE;
			while(file == INVALID_HANDLE_VALUE && ++tries < 10)
			{
				batchFilename = Format(L"%s\\robo%08x.bat", outputPath.c_str(), timeGetTime());
				file = CreateFile(batchFilename.c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
				if(file == INVALID_HANDLE_VALUE)
				{
					if(GetLastError() == ERROR_ALREADY_EXISTS)
					{
						Sleep(1);
					}
					else
					{
						break;
					}
				}
			}
			if(file != INVALID_HANDLE_VALUE)
			{
				// loop through all files scanned from the clipboard when the menu was shown
				bool error = false;
				for(int fileIndex=0; fileIndex<mFiles.size() && !error; ++fileIndex)
				{
					// get file attributes
					std::wstring inputFilename = mFiles[fileIndex];
					DWORD fileAttributes = GetFileAttributes(inputFilename.c_str());
					if(fileAttributes != INVALID_FILE_ATTRIBUTES)
					{
						// split filename into path and name
						std::wstring batchLine;
						std::wstring fullPath;
						std::wstring fullName;

						std::wstring makeDirCommand(L"mkdir");
						std::wstring roboCommand(L"robocopy /NJH /NJS");

						if(SplitPath(inputFilename.c_str(), fullPath, fullName))
						{
							// add commands to the batch file
							bool isCandidate = (fileAttributes & (FILE_ATTRIBUTE_DEVICE | FILE_ATTRIBUTE_OFFLINE)) == 0;
							bool isFolder = (fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && isCandidate;
							if(isFolder)
							{
								batchLine += Format(L"%s \"%s\\%s\"\r\n", makeDirCommand.c_str(), mDestinationPath.c_str(), fullName.c_str());
								batchLine += Format(L"%s \"%s\" \"%s\\%s\" /E\r\n\r\n", roboCommand.c_str(), inputFilename.c_str(), mDestinationPath.c_str(), fullName.c_str());
							}
							else if(isCandidate)
							{
								batchLine += Format(L"%s \"%s\" \"%s\" \"%s\"\r\n\r\n", roboCommand.c_str(), fullPath.c_str(), mDestinationPath.c_str(), fullName.c_str());
							}
							else
							{
								batchLine += Format(L"REM skipped \"%s\"\r\n\r\n", inputFilename.c_str());
							}

							if(!batchLine.empty())
							{
								// make the line ansi
								std::vector<char> mbc = WideStringToAnsi(batchLine);

								// write it to the batch file
								DWORD wrote = 0;
								if(!WriteFile(file, &mbc[0], (DWORD)mbc.size(), &wrote, NULL))
								{
									error = true;
									MessageBox(hWnd, Format(L"Error writing to Robo batch file: %08x", batchFilename.c_str()).c_str(), L"RoboPaste", MB_ICONEXCLAMATION);
									break;
								}
							}
						}
						else
						{
							ShowError(hWnd, L"Error parsing filename %s", inputFilename.c_str());
						}
					}
					else
					{
						switch (MessageBox(hWnd, Format(L"Error getting file information for %s", inputFilename.c_str()).c_str(), L"RoboPaste", MB_ABORTRETRYIGNORE))
						{
						case IDABORT:
							error = true;
							break;
							
						case IDRETRY:
							--fileIndex;
							break;

						case IDIGNORE:
							break;
						};
					}
				}
				CloseHandle(file);

				if(!error)
				{
					switch(CustomMessageBox(hWnd, Format(L"Batch file %s created.", batchFilename.c_str()).c_str(), L"Run it", L"Cancel", L"Show in folder"))
					{
					case IDYES:
						{
							std::wstring params = Format(L"/T:06 /Q /K %s", batchFilename.c_str());
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
								ShowError(hWnd, L"Error executing batch file (%08x)", GetLastError());
							}
						}
						break;

					case IDNO:
						break;

					case IDCANCEL:
						BrowseToFile(batchFilename.c_str());
						break;
					}
				}
			}
			else
			{
				ShowError(hWnd, L"Error creating Robo batch file %s", batchFilename.c_str());
			}
		}
		else
		{
			ShowError(hWnd, L"Error creating RoboBatch folder (%s): %08x", outputPath.c_str(), GetLastError());
		}
	}
	else
	{
		ShowError(hWnd, L"Error finding AppData folder");
	}

}

//////////////////////////////////////////////////////////////////////

#pragma region IUnknown

//////////////////////////////////////////////////////////////////////
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

//////////////////////////////////////////////////////////////////////
// Increase the reference count for an interface on an object.

IFACEMETHODIMP_(ULONG) FileContextMenuExt::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

//////////////////////////////////////////////////////////////////////
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

//////////////////////////////////////////////////////////////////////

#pragma region IShellExtInit

//////////////////////////////////////////////////////////////////////
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

//////////////////////////////////////////////////////////////////////

#pragma endregion

//////////////////////////////////////////////////////////////////////

#pragma region IContextMenu

//////////////////////////////////////////////////////////////////////

IFACEMETHODIMP FileContextMenuExt::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

	// Check the clipboard
	// If it's got some filenames in it
	// Enable the RoboPaste entry

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
			for(UINT i=0; i<nFiles; ++i)
			{
				DragQueryFile((HDROP)stgm.hGlobal, i, path, ARRAYSIZE(path));
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
    mii.hbmpItem = NULL;
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

//////////////////////////////////////////////////////////////////////

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
            OnRoboPaste(pici->hwnd);
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
            OnRoboPaste(pici->hwnd);
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
            OnRoboPaste(pici->hwnd);
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

//////////////////////////////////////////////////////////////////////
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

//////////////////////////////////////////////////////////////////////

#pragma endregion