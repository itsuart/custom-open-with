#include "shell-link.h"
#include <windows.h>
#include <ObjIdl.h>
#include <Shobjidl.h>

#define ODS(x) OutputDebugStringW(x)
//#define ODS(x) (void)0

static u16 temp_file_path[MAX_UNICODE_PATH_LENGTH + 1] = {0};
static u16 target_file_path[MAX_UNICODE_PATH_LENGTH + 1] = {0};

/*
  .lnk file will be saved at temporary location and moved to proper with SHFileOperation because it will create missing directories for us
 */

bool ShellLink_CreateAndStore(u16* linkTargetPath, u16* linkTargetArguments, u16* whereToStore){
    SecureZeroMemory(temp_file_path, sizeof(temp_file_path));
    SecureZeroMemory(target_file_path, sizeof(target_file_path));

    IShellLinkW* pIShellLink = NULL;

    HRESULT hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLinkW, (void**)&pIShellLink);
    if (hr != S_OK) return false;

    ODS(L"Created instance");

    hr = COM_CALL(pIShellLink, SetPath, linkTargetPath);
    if (! SUCCEEDED(hr)){
        COM_CALL0(pIShellLink, Release);
        return false;
    }

    ODS(L"Path is set:"); ODS(linkTargetPath);

    hr = COM_CALL(pIShellLink, SetArguments, linkTargetArguments);
    if (! SUCCEEDED(hr)){
        COM_CALL0(pIShellLink, Release);
        return false;
    }

    ODS(L"ARgs are set:"); ODS(linkTargetArguments);

    IPersistFile* pIPersistFile = NULL;
    hr = COM_CALL(pIShellLink, QueryInterface, &IID_IPersistFile, (void**) &pIPersistFile);
    if (! SUCCEEDED(hr)){
        COM_CALL0(pIShellLink, Release);
        return false;
    }

    {
        const uint max_size = sizeof(temp_file_path) / sizeof(temp_file_path[0]);
        uint pathLength = GetTempPathW(max_size, temp_file_path);
        if (pathLength == 0 || pathLength > max_size){
            COM_CALL0(pIPersistFile, Release);
            COM_CALL0(pIShellLink, Release);
            return false;
        }

        uint result = GetTempFileNameW(temp_file_path, L"cow", 0, temp_file_path);
        if (result == 0){
            COM_CALL0(pIPersistFile, Release);
            COM_CALL0(pIShellLink, Release);
            return false;
        }

        ODS(L"temp_file_path:"); ODS(temp_file_path);
    }


    hr = COM_CALL(pIPersistFile, Save, temp_file_path, true);
    COM_CALL0(pIPersistFile, Release);
    COM_CALL0(pIShellLink, Release);

    if (! SUCCEEDED(hr)){
        DeleteFileW(temp_file_path);
        return false;
    }

    lstrcpyW(target_file_path, whereToStore);

    SHFILEOPSTRUCT fileOp = {
        .hwnd = NULL,
        .wFunc = FO_MOVE,
        .pFrom = temp_file_path,
        .pTo = target_file_path,
        .fFlags = FOF_NO_UI | FOF_RENAMEONCOLLISION,
        .fAnyOperationsAborted = false,
        .hNameMappings = NULL,
        .lpszProgressTitle = L"Saving custom open with configuration"
    };

    bool succeeded = (0 == SHFileOperation(&fileOp));
    succeeded |= fileOp.fAnyOperationsAborted;

    DeleteFileW(temp_file_path);
    SecureZeroMemory(temp_file_path, sizeof(temp_file_path));
    SecureZeroMemory(target_file_path, sizeof(target_file_path));

    ODS(succeeded ? L"Save succeeded" : L"Save failed");
    return succeeded;
}

#undef ODS
