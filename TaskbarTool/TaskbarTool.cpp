// TaskbarTool.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"

using namespace std;

BOOL TaskbarPin(LPTSTR lpFilePath, BOOL bIsPin = FALSE)
{
	BOOL bRet = FALSE;
	HMENU hmenu = NULL;
	LPSHELLFOLDER pdf = NULL;
	LPSHELLFOLDER psf = NULL;
	LPITEMIDLIST pidl = NULL;
	LPITEMIDLIST pitm = NULL;
	LPCONTEXTMENU pcm = NULL;

	TCHAR szFilePath[MAX_PATH] = {0};
	_tcscpy_s(szFilePath, MAX_PATH, lpFilePath);

	TCHAR* pszFileName = PathFindFileName(szFilePath);
	PathRemoveFileSpec(szFilePath);

	wcout << szFilePath << endl;
	wcout << pszFileName << endl;

	if (SUCCEEDED(SHGetDesktopFolder(&pdf))
		&& SUCCEEDED(pdf->ParseDisplayName(NULL, NULL, szFilePath, NULL, &pidl,  NULL))
		&& SUCCEEDED(pdf->BindToObject(pidl, NULL, IID_IShellFolder, (void **)&psf))
		&& SUCCEEDED(psf->ParseDisplayName(NULL, NULL, pszFileName, NULL, &pitm,  NULL))
		&& SUCCEEDED(psf->GetUIObjectOf(NULL, 1, (LPCITEMIDLIST *)&pitm, IID_IContextMenu, NULL, (void **)&pcm))
		&& (hmenu = CreatePopupMenu()) != NULL
		&& SUCCEEDED(pcm->QueryContextMenu(hmenu, 0, 1, INT_MAX, CMF_NORMAL)))
	{
		CMINVOKECOMMANDINFO ici = { sizeof(CMINVOKECOMMANDINFO), 0 };
		ici.hwnd = NULL;
		ici.lpVerb = bIsPin ? "taskbarpin" : "taskbarunpin";
		pcm->InvokeCommand(&ici);
		bRet = TRUE;
	}

	if (hmenu)
		DestroyMenu(hmenu);
	if (pcm)
		pcm->Release();
	if (pitm)
		CoTaskMemFree(pitm);
	if (psf)
		psf->Release();
	if (pidl)
		CoTaskMemFree(pidl);
	if (pdf)
		pdf->Release();

	return bRet;
}

BOOL CreateShortcut(int argc, LPTSTR* argv)
{
	BOOL			bRet = TRUE;
	HRESULT			hRet;
	IShellLink*		pLink;
	IPersistFile*	ppf;

	TCHAR	szFilePath[MAX_PATH]	= { 0 };		// src file path
	TCHAR	szLnkPath[MAX_PATH]		= { 0 };		// lnk file path
	TCHAR	szParams[MAX_PATH]		= { 0 };		// arguments
	TCHAR	szIconPath[MAX_PATH]	= { 0 };		// lnk icon
	TCHAR	szIconIndex[MAX_PATH]	= { 0 };		// TCHAR icon index
	int		nIconIndex				= 0;			// icon index

	_tcscpy_s(szFilePath, argv[2]);
	_tcscpy_s(szLnkPath, argv[3]);
	if (argc >= 4)
	{
		_tcscpy_s(szParams, argv[4]);
	}

	if (argc >= 5)
	{
		_tcscpy_s(szIconPath, argv[5]);
	}

	if (argc >= 6)
	{
		_tcscpy_s(szIconIndex, argv[6]);
		nIconIndex = _ttoi(szIconIndex);
	}

	hRet = ::CoCreateInstance(CLSID_ShellLink, NULL,CLSCTX_INPROC_SERVER, IID_IShellLink, (void**)&pLink);
	if ( hRet != S_OK)
	{
		cout << "Failed to create shortcut！" << endl;
		bRet = FALSE;
	}
	else
	{
		hRet = pLink->QueryInterface(IID_IPersistFile, (void**)&ppf);
		if (hRet != S_OK)
		{
			cout << "Failed to create shortcut！" << endl;
			bRet = FALSE;
		}
		else
		{
			pLink->SetPath(szFilePath);

			if (_tcslen(szParams) > 0)
			{
				pLink->SetArguments(szParams);
			}

			if (_tcslen(szIconPath) > 0)
			{
				pLink->SetIconLocation(szIconPath, nIconIndex);
			}

			ppf->Save(szLnkPath, TRUE);
		}
	}

	ppf->Release();
	pLink->Release();

	return bRet;
}

int _tmain(int argc, _TCHAR* argv[])
{
	if (argc < 3)
	{
		cout << "调用示例，By https://www.xiaobanma.com/" << endl;
		cout << "将快捷方式锁定到任务栏：TaskbarTool /pin C:\\Windows\\notepad.exe" << endl;
		cout << "将快捷方式从任务栏解锁：TaskbarTool /unpin C:\\Windows\\notepad.exe" << endl;
		cout << "记事本创建快捷方式到桌面：TaskbarTool /lnk C:\\Windows\\notepad.exe C:\\Users\\Administrator\\Desktop\\notepad.lnk C:\\1.txt" << endl;
		return -1;
	}

	CoInitialize(nullptr);

	if (_tcsicmp(argv[1], _T("/pin")) == 0)
	{
		wcout << _T("Pin: ") << argv[2] << " " << _T("to taskbar.") << endl;
		TaskbarPin(argv[2], TRUE);
	}
	else if (_tcsicmp(argv[1], _T("/unpin")) == 0)
	{
		wcout << _T("UnPin: ") << argv[2] << " " << _T("from taskbar.") << endl;
		TaskbarPin(argv[2]);
	}
	else if (_tcsicmp(argv[1], _T("/lnk")) == 0 && argc >= 4)
	{
		wcout << _T("CreateLnk: ") << argv[2] << endl;
		CreateShortcut(argc, argv);
	}

	CoUninitialize();

	return 0;
}

