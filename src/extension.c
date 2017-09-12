#define _WIN32_WINNT _WIN32_WINNT_WIN7
#define UNICODE
#define STRICT_TYPED_ITEMIDS    // Better type safety for IDLists

#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>
#include <Shobjidl.h>
#include <Shlobj.h>

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;
typedef unsigned char u8;

//#define ODS(x) OutputDebugStringW(x)
#define ODS(x) (void)0

#define COM_CALL(x, y, ...) x->lpVtbl->y(x, __VA_ARGS__)
#define COM_CALL0(x, y) x->lpVtbl->y(x)

/* GLOBALS */
#include "common.c"

/* FORWARD DECLARATION EVERYTHING */
struct tag_IMyObj;

ULONG STDMETHODCALLTYPE myObj_AddRef(struct tag_IMyObj* pMyObj);
ULONG STDMETHODCALLTYPE myObj_Release(struct tag_IMyObj* pMyObj);
HRESULT STDMETHODCALLTYPE myObj_QueryInterface(struct tag_IMyObj* pMyObj, REFIID requestedIID, void **ppv);

/* ALL THE STRUCTS (because it's too much pain to have them in separate places) */
typedef struct tag_MyIContextMenuImpl {
    IContextMenuVtbl* lpVtbl;
    struct tag_IMyObj* pBase;
} MyIContextMenuImpl;

typedef struct tag_MyIShellExtInitImpl {
    IShellExtInitVtbl* lpVtbl;
    struct tag_IMyObj* pBase;
} MyIShellExtInitImpl;

typedef struct tag_IMyObj{
    IUnknownVtbl* lpVtbl;
    MyIContextMenuImpl contextMenuImpl;
    MyIShellExtInitImpl shellExtInitImpl;
    uint nItems;
    u16* itemNames;
    long int refsCount;
} IMyObj;

static void cleanup_names_storage(IMyObj* pBase){
    if (pBase->itemNames != NULL){
        GlobalFree(pBase->itemNames);
    }
    pBase->itemNames = NULL;
    pBase->nItems = 0;
}



/******************************* IContextMenu ***************************/
static const u16* COPY_VERB = SERVER_GUID_TEXT L".COPY";
static const u16* MOVE_VERB = SERVER_GUID_TEXT L".MOVE";

#define COPY_TO_MENU_OFFSET 1
#define MOVE_TO_MENU_OFFSET 2

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_GetCommandString(MyIContextMenuImpl* pImpl,
                                                              UINT_PTR idCmd, UINT uFlags, UINT* pwReserved,LPSTR pszName,UINT cchMax){
    if (uFlags == GCS_VERBW){
        switch (idCmd){
            case COPY_TO_MENU_OFFSET:{
                lstrcpyW((u16*)pszName, COPY_VERB);
                ODS(L"GetCommandString SUCCESS (copy verb)");
                return S_OK;
            } break;

            case MOVE_TO_MENU_OFFSET:{
                lstrcpyW((u16*)pszName, MOVE_VERB);
                ODS(L"GetCommandString SUCCESS (move verb)");
                return S_OK;
            } break;

        }
    } else if (uFlags == GCS_VALIDATEW){
        if (idCmd == COPY_TO_MENU_OFFSET || idCmd == MOVE_TO_MENU_OFFSET){
            ODS(L"GetCommandString SUCCESS (validate cmd)");
            return S_OK;
        } else {
            ODS(L"GetCommandString UNKNOWN CMD VALIDATION");
            return S_FALSE;
        }
    }
    ODS(L"GetCommandString FAILED");
    return E_INVALIDARG;
}

//returns number of digits in the number
static uint u64toW(uint number, u16 stringBuffer[20]){
    const static u16 DIGITS[10] = L"0123456789";
    if (number == 0){ //special case
        stringBuffer[0] = L'0';
        return 1;
    }

    //biggest 64 bit unsigned number is 18,446,744,073,709,551,615 that is 20 digits, so
    uint divisor = 10;
    int power = 0;
    u8 reminders[20] = {0};

    for (; power < 20; power += 1){
        char reminder = number % divisor;

        reminders[power] = reminder;
        number = number / divisor;

        if (number == 0){
            break;
        }
    }

    //put 'em digits in correct order
    for (int i = power; i >= 0; i -= 1){
        stringBuffer[power - i] = DIGITS[reminders[i]];
    }

    return power;
}

static void mk_label (u16* prefix, uint number, u16* buffer){
    uint index = 0;
    while (prefix[index] != 0){
        buffer[index] = prefix[index];
        index += 1;
    }
    uint nDigits = u64toW(number, buffer + index);
    index += nDigits;

    lstrcpyW(buffer + index + 1, number > 1 ? L" items to..." : L" item to...");
}

static const DWORD ID_OPEN_TARGET_DIRECTORY = 1000;

static bool query_user_for_target_file_path(HWND ownerWnd, u16* sourceFilePath, u16* pTargetPath, int* pOpenTargetFolder){
    *pOpenTargetFolder = false;

    IShellItem* sourceItem = NULL;
    if (! SUCCEEDED(SHCreateItemFromParsingName(sourceFilePath, NULL, &IID_IShellItem, (void**)&sourceItem))){
        return false;
    }

    bool succeeded = false;
    HRESULT hr = S_OK;
    // Create a new common save file dialog.
    IFileSaveDialog *pfd = NULL;

    hr = CoCreateInstance(&CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileSaveDialog, (void**)&pfd);
    if (SUCCEEDED(hr)){
        {
            DWORD dwOptions;
            hr = COM_CALL(pfd, GetOptions, &dwOptions);
            if (SUCCEEDED(hr)){
                //Don't ask 'are you sure you want to overwrite' at this moment - SHFileOperation is better at it.
                hr = COM_CALL(pfd, SetOptions, (dwOptions | FOS_CREATEPROMPT) & (~FOS_OVERWRITEPROMPT));
            }
        }

        hr = COM_CALL(pfd, SetSaveAsItem, sourceItem);

        IFileDialogCustomize* customizer = NULL;
        hr = COM_CALL(pfd, QueryInterface, &IID_IFileDialogCustomize, (void**)&customizer);
        if (SUCCEEDED(hr)){
            COM_CALL(customizer, AddCheckButton, ID_OPEN_TARGET_DIRECTORY, L"Open target folder", FALSE);

            // Show the save file dialog.
            if (SUCCEEDED(hr)){
                hr = COM_CALL(pfd, Show, ownerWnd);
                if (SUCCEEDED(hr)){

                    COM_CALL(customizer, GetCheckButtonState, ID_OPEN_TARGET_DIRECTORY, pOpenTargetFolder);

                    // Get the selection from the user.
                    IShellItem* psiResult = NULL;
                    hr = COM_CALL(pfd, GetResult, &psiResult);
                    if (SUCCEEDED(hr)){
                        u16* pszPath = NULL;
                        hr = COM_CALL(psiResult, GetDisplayName, SIGDN_FILESYSPATH, &pszPath);

                        if (SUCCEEDED(hr)){
                            lstrcpyW(pTargetPath, pszPath);
                            CoTaskMemFree(pszPath);

                            succeeded = true;

                        }
                        COM_CALL0(psiResult, Release);
                    }
                }
            }
            COM_CALL0(customizer, Release);
        }
        COM_CALL0(pfd, Release);
    }
    COM_CALL0(sourceItem, Release);

    return succeeded;
}
static bool query_user_for_target_folder(HWND ownerWnd, u16* title, u16* startFolder, u16* pTargetFolder, int* pOpenTargetFolder){
    bool succeeded = false;
    *pOpenTargetFolder = false;

    HRESULT hr = S_OK;
    // Create a new common open file dialog.
    IFileOpenDialog* pfd = NULL;
    hr = CoCreateInstance(&CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileOpenDialog, (void**)&pfd);
    if (SUCCEEDED(hr)){
        {
            // Set the dialog as a folder picker.
            DWORD dwOptions;
            hr = COM_CALL(pfd, GetOptions, &dwOptions);
            if (SUCCEEDED(hr)){
                hr = COM_CALL(pfd, SetOptions, dwOptions | FOS_PICKFOLDERS | FOS_CREATEPROMPT);
            }
        }

        IFileDialogCustomize* customizer = NULL;
        hr = COM_CALL(pfd, QueryInterface, &IID_IFileDialogCustomize, (void**)&customizer);
        if (SUCCEEDED(hr)){
            COM_CALL(customizer, AddCheckButton, ID_OPEN_TARGET_DIRECTORY, L"Open target folder", FALSE);

            // Set the title of the dialog.
            if (SUCCEEDED(hr)){
                hr = COM_CALL(pfd, SetTitle, title);

                if (NULL != startFolder){
                    IShellItem* startShellItem;
                    if (SUCCEEDED(SHCreateItemFromParsingName(startFolder, NULL, &IID_IShellItem, (void**)&startShellItem))){
                        COM_CALL(pfd, SetFolder, startShellItem);
                        COM_CALL0(startShellItem, Release);
                    } //we don't really care if we fail to set starting folder
                }

                // Show the folder picker.
                if (SUCCEEDED(hr)){
                    hr = COM_CALL(pfd, Show, ownerWnd);
                    if (SUCCEEDED(hr)){
                        COM_CALL(customizer, GetCheckButtonState, ID_OPEN_TARGET_DIRECTORY, pOpenTargetFolder);

                        // Get the selection from the user.
                        IShellItem* psiResult = NULL;
                        hr = COM_CALL(pfd, GetResult, &psiResult);
                        if (SUCCEEDED(hr)){
                            u16* pszPath = NULL;
                            hr = COM_CALL(psiResult, GetDisplayName, SIGDN_FILESYSPATH, &pszPath);
                            if (SUCCEEDED(hr)){
                                lstrcpyW(pTargetFolder, pszPath);
                                CoTaskMemFree(pszPath);

                                succeeded = true;

                            }

                            COM_CALL0(psiResult, Release);
                        }
                    }
                }
            }
        }


        COM_CALL0(customizer, Release);
        COM_CALL0(pfd, Release);
    }

    return succeeded;
}

//I really hate person that causes me "undefined reference to `___chkstk_ms'" errors. And extra calls to clear the memory.
static u16 targetDir[MAX_UNICODE_PATH_LENGTH + 2]; //extra 1 for trailing zero
static u16 startDirContainer[MAX_UNICODE_PATH_LENGTH + 1];

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_InvokeCommand(MyIContextMenuImpl* pImpl, LPCMINVOKECOMMANDINFO pCommandInfo){
    IMyObj* pBase = pImpl->pBase;

    if (pBase->nItems == 0){
        return S_OK; //instantly copy or move nothing!
    }

    const void* pVerb = pCommandInfo->lpVerb;

    if (HIWORD(pVerb)){
        return E_INVALIDARG;
    }

    u16 title[(sizeof(L"COPY ")/sizeof(L' ')) + 20 + (sizeof(L" items to...")/sizeof(L' ')) + 1];
    SecureZeroMemory(&title[0], sizeof(title));
    uint fileOp;

    if (pVerb == (void*)MAKEINTRESOURCE(COPY_TO_MENU_OFFSET)){
        mk_label(L"COPY ", pBase->nItems, &title[0]);
        fileOp = FO_COPY;
    } else if (pVerb == (void*)MAKEINTRESOURCE(MOVE_TO_MENU_OFFSET)){
        mk_label(L"MOVE ", pBase->nItems, &title[0]);
        fileOp = FO_MOVE;
    } else {
        ODS(L"Neither COPY nor MOVE offset");
        cleanup_names_storage(pBase);
        return E_FAIL;
    }

    //NOTE: we assuming here that name of first item does NOT end with a BACKSLASH
    uint lastSlashPos = 0;
    startDirContainer[sizeof(startDirContainer) / sizeof(startDirContainer[0]) - 1] = 0;
    for (uint i = 0; i < MAX_UNICODE_PATH_LENGTH; i += 1){
        u16 ch = pBase->itemNames[i];

        if (ch == L'\\') {
            lastSlashPos = i;
        }

        startDirContainer[i] = ch;

        if (ch == 0) {
            startDirContainer[lastSlashPos] = 0;
            break;
        }
    }

    SecureZeroMemory(targetDir, sizeof(targetDir));

    const bool operating_on_single_file = pBase->nItems == 1 && (0 == (FILE_ATTRIBUTE_DIRECTORY & GetFileAttributesW(pBase->itemNames)));
    int open_target_folder = 0;
    bool target_directory_acquired  = false;

    //Alright let us see if we have single file to process (actually we have to check that it is NOT DIRECTORY)
    if (operating_on_single_file){
        target_directory_acquired = query_user_for_target_file_path(pCommandInfo->hwnd, pBase->itemNames, &targetDir[0], &open_target_folder);
    } else {
        target_directory_acquired = query_user_for_target_folder(pCommandInfo->hwnd, title, startDirContainer, &targetDir[0], &open_target_folder);
    }


    if (target_directory_acquired){
        SHFILEOPSTRUCT shFileOp = {
            .hwnd = pCommandInfo->hwnd,
            .wFunc = fileOp,
            .pFrom = pBase->itemNames,
            .pTo = &targetDir[0],
            .fFlags = FOF_ALLOWUNDO,
            .fAnyOperationsAborted = false,
            .hNameMappings = NULL,
            .lpszProgressTitle = L"ERROR!"
        };

        if ( (0 == SHFileOperationW(&shFileOp)) && (! shFileOp.fAnyOperationsAborted) && open_target_folder){
            if (operating_on_single_file){
                PIDLIST_ABSOLUTE path = ILCreateFromPathW(targetDir); //I get it now - it will return null if file doesn't exist. That is why it works sometime.
                SHOpenFolderAndSelectItems(path, 0, NULL, 0);
                ILFree(path);
            } else {
                ShellExecuteW(NULL, L"explore", targetDir, NULL, NULL, SW_SHOWNORMAL);
            }
        }
    }

    cleanup_names_storage(pBase);
    return S_OK;
}


HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryContextMenu(MyIContextMenuImpl* pImpl,
                                                           HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags){
    IMyObj* pBase = pImpl->pBase;

    if (CMF_DEFAULTONLY == (CMF_DEFAULTONLY & uFlags)
        || CMF_VERBSONLY == (CMF_VERBSONLY & uFlags)
        || CMF_NOVERBS == (CMF_NOVERBS & uFlags)){

        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //in highly unlikely case where is no room left in the menu:
    if (idCmdFirst + 2 > idCmdLast){
        //we don't add any menu enries of ours
        cleanup_names_storage(pBase);
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    //insert our menu items: separator, copy and move in that order.
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + COPY_TO_MENU_OFFSET, L"Copy To...");
    InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + MOVE_TO_MENU_OFFSET, L"Move To...");

    ODS(L"QueryContextMenu");

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, MOVE_TO_MENU_OFFSET + 1); //shouldn't this be +idCmdFirst?
}


ULONG STDMETHODCALLTYPE myIContextMenuImpl_AddRef(MyIContextMenuImpl* pImpl){
    return myObj_AddRef(pImpl->pBase);
}

ULONG STDMETHODCALLTYPE myIContextMenuImpl_Release(MyIContextMenuImpl* pImpl){
    return myObj_Release(pImpl->pBase);
}

HRESULT STDMETHODCALLTYPE myIContextMenuImpl_QueryInterface(MyIContextMenuImpl* pImpl, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pImpl;
        myObj_AddRef(pImpl->pBase);
        return S_OK;
    }

    return myObj_QueryInterface(pImpl->pBase, requestedIID, ppv);

}

static IContextMenuVtbl IMyIContextMenuVtbl = {
    .QueryInterface = &myIContextMenuImpl_QueryInterface,
    .AddRef = &myIContextMenuImpl_AddRef,
    .Release = &myIContextMenuImpl_Release,
    .GetCommandString = &myIContextMenuImpl_GetCommandString,
    .QueryContextMenu = &myIContextMenuImpl_QueryContextMenu,
    .InvokeCommand = &myIContextMenuImpl_InvokeCommand
};



/************************ IShellExtInit **********************************/
HRESULT STDMETHODCALLTYPE myIShellExtInit_AddRef(MyIShellExtInitImpl* pImpl){
    return myObj_AddRef(pImpl->pBase);
}

ULONG STDMETHODCALLTYPE myIShellExtInit_Release(MyIShellExtInitImpl* pImpl){
    return myObj_Release(pImpl->pBase);
}

HRESULT STDMETHODCALLTYPE myIShellExtInit_QueryInterface(MyIShellExtInitImpl* pImpl, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pImpl;
        myObj_AddRef(pImpl->pBase);
        return S_OK;
    }

    return myObj_QueryInterface(pImpl->pBase, requestedIID, ppv);

}

HRESULT STDMETHODCALLTYPE myIShellExtInit_Initialize(MyIShellExtInitImpl* pImpl, LPCITEMIDLIST pIDFolder, IDataObject* pDataObj, HKEY hRegKey){
    FORMATETC   fe = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM   medium;

    if (pImpl->pBase->itemNames != NULL){
        GlobalFree(pImpl->pBase->itemNames);
        pImpl->pBase->itemNames = NULL;
        pImpl->pBase->nItems = 0;
    }

    if (SUCCEEDED(pDataObj->lpVtbl->GetData(pDataObj, &fe, &medium))){
        // Get the count of files dropped.
        uint uCount = DragQueryFile((HDROP)medium.hGlobal, (UINT)-1, NULL, 0);
        if (uCount > 0){
            u16* allNamesBuffer = GlobalAlloc(GPTR, (uCount * (MAX_UNICODE_PATH_LENGTH + 1) + 1) * sizeof(u16));
            if (allNamesBuffer == NULL){
                ReleaseStgMedium(&medium);
                return E_OUTOFMEMORY;
            }

            //copy file items to the buffer
            u16* currentName = allNamesBuffer;
            for (uint i = 0; i < uCount; i += 1){
                uint nCharsCopied = DragQueryFileW((HDROP)medium.hGlobal, i, currentName, MAX_UNICODE_PATH_LENGTH);
                if (nCharsCopied){
                    currentName += (nCharsCopied + 1);
                } else {/*copying fucked up */}
            }
            pImpl->pBase->nItems = uCount;
            pImpl->pBase->itemNames = allNamesBuffer;
        }

        ReleaseStgMedium(&medium);
    }

    ODS(L"Initialize success");

    return S_OK;
}

static IShellExtInitVtbl IMyIShellExtVtbl = {
    .AddRef = &myIShellExtInit_AddRef,
    .Release = &myIShellExtInit_Release,
    .QueryInterface = &myIShellExtInit_QueryInterface,
    .Initialize = &myIShellExtInit_Initialize
};



/******************* IMyObj  *******************************/
static long int nObjectsAndRefs = 0;

ULONG STDMETHODCALLTYPE myObj_AddRef(IMyObj* pMyObj){
    InterlockedIncrement(&nObjectsAndRefs);

    InterlockedIncrement(&(pMyObj->refsCount));
    return pMyObj->refsCount;
}

ULONG STDMETHODCALLTYPE myObj_Release(IMyObj* pMyObj){
    InterlockedDecrement(&nObjectsAndRefs);

    InterlockedDecrement(&(pMyObj->refsCount));
    long int nRefs = pMyObj->refsCount;

    if (nRefs == 0){
        if (pMyObj->itemNames != NULL){
            GlobalFree(pMyObj->itemNames);
            pMyObj->itemNames = NULL;
            pMyObj->nItems = 0;
        }
        GlobalFree(pMyObj);
    }

    return nRefs;
}

HRESULT STDMETHODCALLTYPE myObj_QueryInterface(IMyObj* pMyObj, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }
    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pMyObj;
        myObj_AddRef(pMyObj);
        return S_OK;
    }

    if (IsEqualGUID(requestedIID, &IID_IContextMenu)){
        *ppv = &(pMyObj->contextMenuImpl);
        myObj_AddRef(pMyObj);

        return S_OK;
    }

    if (IsEqualGUID(requestedIID, &IID_IShellExtInit)){
        *ppv = &(pMyObj->shellExtInitImpl);
        myObj_AddRef(pMyObj);

        return S_OK;
    }


    return E_NOINTERFACE;
}

static IUnknownVtbl IMyObjVtbl = {
    .QueryInterface = &myObj_QueryInterface,
    .AddRef = &myObj_AddRef,
    .Release = &myObj_Release
};



/************** IClassFactory *****************************/
HRESULT STDMETHODCALLTYPE classCreateInstance(IClassFactory* pClassFactory, IUnknown* punkOuter, REFIID pRequestedIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

   // Assume an error by clearing caller's handle.
   *ppv = 0;

   // We don't support aggregation in IExample.
   if (punkOuter){
       return CLASS_E_NOAGGREGATION;
   }

   if (IsEqualGUID(pRequestedIID, &IID_IUnknown)
       || IsEqualGUID(pRequestedIID, &IID_IContextMenu)
       || IsEqualGUID(pRequestedIID, &IID_IShellExtInit)
       ){
       IMyObj* pMyObj = GlobalAlloc(GPTR, sizeof(IMyObj)); //Zero the memory
       if (pMyObj == NULL){
           return E_OUTOFMEMORY;
       }

       pMyObj->lpVtbl = &IMyObjVtbl;

       pMyObj->contextMenuImpl.lpVtbl = &IMyIContextMenuVtbl;
       pMyObj->contextMenuImpl.pBase = pMyObj;

       pMyObj->shellExtInitImpl.lpVtbl = &IMyIShellExtVtbl;
       pMyObj->shellExtInitImpl.pBase = pMyObj;

       return myObj_QueryInterface(pMyObj, pRequestedIID, ppv);
   }

   return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE classAddRef(IClassFactory* pClassFactory){
   return 1;
}

ULONG STDMETHODCALLTYPE classRelease(IClassFactory* pClassFactory){
   return 1;
}

long int nLocks = 0;

HRESULT STDMETHODCALLTYPE classLockServer(IClassFactory* pClassFactory, BOOL flock){
    if (flock){
        InterlockedIncrement(&nLocks);
    } else {
        InterlockedDecrement(&nLocks);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE classQueryInterface(IClassFactory* pClassFactory, REFIID requestedIID, void **ppv){
   // Check if the GUID matches an IClassFactory or IUnknown GUID.
   if (! IsEqualGUID(requestedIID, &IID_IUnknown) &&
       ! IsEqualGUID(requestedIID, &IID_IClassFactory)){
      // It doesn't. Clear his handle, and return E_NOINTERFACE.
      *ppv = 0;
      return E_NOINTERFACE;
   }

   *ppv = pClassFactory;
    pClassFactory->lpVtbl->AddRef(pClassFactory);
   return S_OK;
}

static const IClassFactoryVtbl IClassFactory_Vtbl = {
    classQueryInterface,
    classAddRef,
    classRelease,
    classCreateInstance,
    classLockServer
};

//since we need only one ClassFactory ever:
static IClassFactory IClassFactoryObj = {
    .lpVtbl = &IClassFactory_Vtbl
};


/************************** DLL STUFF ****************************************/

/*
  COM subsystem will call this function in attempt to retrieve an implementation of an IID.
  Usually it's IClassFactory that is requested (but could be anything else?)
 */
__declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID pCLSID, REFIID pIID, void** ppv){
    if (ppv == NULL){
        return E_POINTER;
    }

    *ppv = NULL;
    if (! IsEqualGUID(pCLSID, &SERVER_GUID)){
        //that's not our CLSID, ignore
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    if (! IsEqualGUID(pIID, &IID_IClassFactory)){
        //we only providing ClassFactories in this function
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    //alright, return our implementation of IClassFactory
    *ppv = &IClassFactoryObj;
    return S_OK;
}

__declspec(dllexport) HRESULT PASCAL DllCanUnloadNow(){
    if (nLocks == 0 && nObjectsAndRefs == 0){
        return S_OK;
    }
    return S_FALSE;
}

bool WINAPI entry_point(HINSTANCE hInst, DWORD reason, void* pReserved){
    return true;
}
