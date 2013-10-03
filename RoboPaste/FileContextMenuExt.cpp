//////////////////////////////////////////////////////////////////////

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <Mmsystem.h>
#include <ctime>
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "winmm.lib")
#include "Reg.h"

//////////////////////////////////////////////////////////////////////

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

static HHOOK customMessageBoxHook;

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

static std::wstring CustomMessageBoxStrings[3] =
{
	std::wstring(L"Yes"),
	std::wstring(L"No"),
	std::wstring(L"Cancel")
};

//////////////////////////////////////////////////////////////////////

LRESULT CALLBACK CustomMessageBoxHookProc(INT nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HCBT_ACTIVATE)
	{
		HWND h = (HWND)wParam;
		SetDlgItemText(h, IDYES, CustomMessageBoxStrings[0].c_str());
		SetDlgItemText(h, IDNO, CustomMessageBoxStrings[1].c_str());
		SetDlgItemText(h, IDCANCEL, CustomMessageBoxStrings[2].c_str());
		UnhookWindowsHookEx(customMessageBoxHook);
	}
	else
	{
		CallNextHookEx(customMessageBoxHook, nCode, wParam, lParam);
	}
	return 0;
}

//////////////////////////////////////////////////////////////////////

static int CustomMessageBox(HWND hwnd, WCHAR const *msg, WCHAR const *yes, WCHAR const *no, WCHAR const *cancel = NULL)
{
	CustomMessageBoxStrings[0] = yes;
	CustomMessageBoxStrings[1] = no;
	CustomMessageBoxStrings[2] = (cancel == NULL) ? L"" : cancel;
	customMessageBoxHook = SetWindowsHookEx(WH_CBT, &CustomMessageBoxHookProc, 0, GetCurrentThreadId());
	return MessageBox(hwnd, msg, L"RoboPaste", (cancel == NULL) ? MB_YESNO : MB_YESNOCANCEL);
}

//////////////////////////////////////////////////////////////////////

static std::wstring Format(WCHAR const *fmt, ...)
{
	WCHAR buf[1024];
	va_list v;
	va_start(v, fmt);
	_vsnwprintf_s(buf, ARRAYSIZE(buf), fmt, v);
	return std::wstring(buf);
}

//////////////////////////////////////////////////////////////////////

static bool SplitPath(WCHAR const *fullPath, std::wstring &path, std::wstring &fileName)
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
	std::wstring fileStr(fnm);
	std::wstring extStr(ext);

	if(pathStr[pathStr.size() - 1] == L'\\')
	{
		pathStr.pop_back();
	}

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

std::string WideStringToAnsiString(std::wstring str)
{
	std::string s;
	int nch = WideCharToMultiByte(CP_ACP, 0, str.c_str(), (int)str.size(), NULL, 0, NULL, NULL);
	if(nch < 32767)	// sanity
	{
		std::vector<char> p((size_t)nch);
		WideCharToMultiByte(CP_ACP, 0, str.c_str(), (int)str.size(), &p[0], nch, NULL, NULL);
		p.push_back(0);
		s = std::string(&p[0]);
	}
	return s;
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

std::string GetDateTime()
{
	time_t rawtime;
	time(&rawtime);

	tm timeinfo;
	localtime_s(&timeinfo, &rawtime);

	char buffer[80];
	strftime(buffer, ARRAYSIZE(buffer), "%Y-%m-%d-%H-%M-%S", &timeinfo);
	return buffer;
}

//////////////////////////////////////////////////////////////////////
// build the batch file as strings

std::vector<std::string> FileContextMenuExt::ScanFiles(HWND hWnd, std::wstring mkdirCommand, std::wstring robocopyCommand)
{
	std::vector<std::string> miscLines;
	std::vector<std::string> mkdirLines;
	std::vector<std::string> folderLines;
	std::vector<std::string> fileLines;

	std::wstring currentFileLine;
	int fileCount = 0;

	miscLines.push_back(std::string("REM RoboPaste batch file created ") + GetDateTime() + "\r\n");

	bool error = false;

	for(size_t fileIndex=0; fileIndex<mFiles.size() && !error; ++fileIndex)
	{
		std::wstring inputFilename = mFiles[fileIndex];
		DWORD fileAttributes = GetFileAttributes(inputFilename.c_str());
		if(fileAttributes != INVALID_FILE_ATTRIBUTES)
		{
			// split filename into path and name
			std::wstring fullPath;
			std::wstring fullName;

			if(SplitPath(inputFilename.c_str(), fullPath, fullName))
			{
				// add commands to the batch file
				bool isCandidate = (fileAttributes & (FILE_ATTRIBUTE_DEVICE | FILE_ATTRIBUTE_OFFLINE)) == 0;
				bool isFolder = (fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && isCandidate;
				if(isFolder)
				{
					mkdirLines.push_back(WideStringToAnsiString(Format(L"%s \"%s\\%s\"\r\n", mkdirCommand.c_str(), mDestinationPath.c_str(), fullName.c_str())));
					folderLines.push_back(WideStringToAnsiString(Format(L"%s \"%s\" \"%s\\%s\" /E\r\n", robocopyCommand.c_str(), inputFilename.c_str(), mDestinationPath.c_str(), fullName.c_str())));
				}
				else if(isCandidate)
				{
					if(currentFileLine.empty())
					{
						currentFileLine = Format(L"%s \"%s\" \"%s\" \"%s\"", robocopyCommand.c_str(), fullPath.c_str(), mDestinationPath.c_str(), fullName.c_str());
					}
					else
					{
						currentFileLine += std::wstring(L" \"") + fullName.c_str() + L"\"";
					}
					++fileCount;
					if((fileCount % 8) == 0)
					{
						currentFileLine += L"\r\n";
						fileLines.push_back(WideStringToAnsiString(currentFileLine));
						currentFileLine.clear();
					}
				}
				else
				{
					miscLines.push_back(WideStringToAnsiString(Format(L"REM skipped \"%s\"\r\n", inputFilename.c_str())));
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
	if(!currentFileLine.empty())
	{
		fileLines.push_back(WideStringToAnsiString(currentFileLine));
	}

	std::vector<std::string> allLines;

	allLines.insert(allLines.end(), miscLines.begin(), miscLines.end());
	allLines.push_back("\r\n");

	allLines.insert(allLines.end(), mkdirLines.begin(), mkdirLines.end());
	allLines.push_back("\r\n");

	allLines.insert(allLines.end(), folderLines.begin(), folderLines.end());
	allLines.push_back("\r\n");

	allLines.insert(allLines.end(), fileLines.begin(), fileLines.end());

	return allLines;
}

//////////////////////////////////////////////////////////////////////

bool FileContextMenuExt::WriteBatchFile(HWND hWnd, std::vector<std::string> lines, std::wstring &batchFilename)
{
	bool rc = false;

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
				// save the lines into the batch file
				bool error = false;
				for(auto line: lines)
				{
					// write it to the batch file
					DWORD wrote = 0;
					if(!WriteFile(file, line.c_str(), (DWORD)line.size(), &wrote, NULL))
					{
						ShowError(hWnd, L"Error writing to Robo batch file: %08x", batchFilename.c_str());
						error = true;
						break;
					}
				}
				CloseHandle(file);
				rc = !error;
			}
			else
			{
				ShowError(hWnd, L"Error creating Robo batch file %s", batchFilename.c_str());
			}
		}
		else
		{
			ShowError(hWnd, L"Error creating Robo batch folder %s", outputPath.c_str());
		}
	}
	else
	{
		ShowError(hWnd, L"Error finding AppData folder");
	}
	return rc;
}

//////////////////////////////////////////////////////////////////////

void FileContextMenuExt::ExecuteBatchFile(HWND hWnd, std::wstring batchFilename)
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

//////////////////////////////////////////////////////////////////////

void FileContextMenuExt::OnRoboPaste(HWND hWnd)
{
	std::wstring makeDirCommand(L"mkdir");
	std::wstring roboCommand = std::wstring(L"robocopy ") + GetRegistryValue(L"parameters", L"/NJH /NJS /MT");

	std::vector<std::string> lines = ScanFiles(hWnd, makeDirCommand, roboCommand);

	if(!lines.empty())
	{
		std::wstring batchFilename;
		if(WriteBatchFile(hWnd, lines, batchFilename))
		{
			switch(CustomMessageBox(hWnd, Format(L"Batch file %s created.", batchFilename.c_str()).c_str(), L"Run it", L"Cancel", L"Show in folder"))
			{
			case IDYES:
				ExecuteBatchFile(hWnd, batchFilename);
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
		ShowError(hWnd, L"No valid files/folders to transfer");
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
		if(mDestinationPath.empty())
		{
			return ERROR_BAD_PATHNAME;
		}
		else
		{
			return S_OK;
		}
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