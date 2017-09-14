#define UNICODE
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>
#include <winerror.h>
#include <Shobjidl.h>
#include <Shlobj.h>
/* GLOBALS */
#include "common.c"

#define ODS(x) OutputDebugStringW(x)

static const DWORD CUSTOM_ID_PARAMS_VISUAL_GROUP = 1000;
static const DWORD CUSTOM_ID_PARAMS_EDITBOX = 1001;

//static const DWROD CUSTOM_ID_SAVE_CONFIGURATION_VISUAL_GROUP = 2000;
static const DWORD CUSTOM_ID_SAVE_CONFIGURATION_RECORD_CHECKBOX = 2001;

//Returns pointer at IFileDialogCustomize if all customizations were applied successfuly, NULL if at least one failed.
static IFileDialogCustomize* customize_dialog(const IFileOpenDialog* pfd, u16 params[MAX_UNICODE_PATH_LENGTH], bool* pSaveConfiguration){
    IFileDialogCustomize* customizer = NULL;
    HRESULT hr = COM_CALL(pfd, QueryInterface, &IID_IFileDialogCustomize, (void**)&customizer);
    if (! SUCCEEDED(hr)) return NULL;

    //Command line parameters controls

    hr = COM_CALL(customizer, StartVisualGroup, CUSTOM_ID_PARAMS_VISUAL_GROUP, L"Command line parameters:");
    if (! SUCCEEDED(hr)) goto fail;

    hr = COM_CALL(customizer, AddEditBox, CUSTOM_ID_PARAMS_EDITBOX, L"Here be params");
    if (! SUCCEEDED(hr)) goto fail;

    hr = COM_CALL0(customizer, EndVisualGroup);
    if (! SUCCEEDED(hr)) goto fail;


    // Store record controls


    hr = COM_CALL(customizer, AddCheckButton, CUSTOM_ID_SAVE_CONFIGURATION_RECORD_CHECKBOX, L"Save this configuration", *pSaveConfiguration);
    if (! SUCCEEDED(hr)) goto fail;

    return customizer;

 fail:
    COM_CALL0(customizer, Release);
    return NULL;
}

static bool get_customized_data(const IFileDialogCustomize* pCustomizer, u16 params[MAX_UNICODE_PATH_LENGTH], bool* pSaveConfiguration){
    HRESULT hr;

    //get the flag wether or not to save the configuration
    hr = COM_CALL(pCustomizer, GetCheckButtonState, CUSTOM_ID_SAVE_CONFIGURATION_RECORD_CHECKBOX, (int*)pSaveConfiguration);
    if (! SUCCEEDED(hr)) return false;

    // get the params text if any
    {
        u16* paramsText = NULL;
        hr = COM_CALL(pCustomizer, GetEditBoxText, CUSTOM_ID_PARAMS_EDITBOX, &paramsText);
        if (! SUCCEEDED(hr)) return false;

        lstrcpyW(params, paramsText);
        CoTaskMemFree(paramsText);
    }
    return true;
}

 typedef struct tag_MyIFileDialogEvents {
     IFileDialogEventsVtbl* lpVtbl;
     IFileDialogCustomize* pCustomizer;
     u16* params;
     bool* pSaveConfiguration;
 } MyIFileDialogEvents;

static HRESULT IFileDialog_OnFileOk(MyIFileDialogEvents* pMyEventHandler, IFileDialog* dontCare){
    //So frustrating :(
    get_customized_data(pMyEventHandler->pCustomizer, pMyEventHandler->params, pMyEventHandler->pSaveConfiguration);
    return S_OK;
}

static HRESULT IFileDialog_OnFolderChange(MyIFileDialogEvents* pMyEventHandlers, IFileDialog* dontCare){ return E_NOTIMPL;}
static HRESULT IFileDialog_OnFolderChanging(MyIFileDialogEvents* pMyEventHandler, IFileDialog* pfd, IShellItem* psiFolder) { return E_NOTIMPL;}
static HRESULT IFileDialog_OnOverwrite(MyIFileDialogEvents* pMyEventHandler, IFileDialog* pfd, IShellItem* psi, FDE_SHAREVIOLATION_RESPONSE* pResponse ) {
    return E_NOTIMPL;
}
static HRESULT IFileDialog_OnSelectionChange(MyIFileDialogEvents* pMyEventHandler, IFileDialog* pfd) {return E_NOTIMPL;}
static HRESULT IFileDialog_OnShareViolation(MyIFileDialogEvents* pMyEventHandler, IFileDialog* pfd, IShellItem* psi, FDE_SHAREVIOLATION_RESPONSE* pResponse) {
    return E_NOTIMPL;
}
static HRESULT IFileDialog_OnTypeChange(MyIFileDialogEvents* pMyEventHandler, IFileDialog* pfd) { return E_NOTIMPL; }

ULONG STDMETHODCALLTYPE IFileDialogEvents_AddRef(void* pMyObj){
    return 1;
}

ULONG STDMETHODCALLTYPE IFileDialogEvents_Release(void* pMyObj){
    return 1;
}

HRESULT STDMETHODCALLTYPE IFileDialogEvents_QueryInterface(void* pMyObj, REFIID requestedIID, void **ppv){
    if (ppv == NULL){
        return E_POINTER;
    }
    *ppv = NULL;

    if (IsEqualGUID(requestedIID, &IID_IFileDialogEvents) || IsEqualGUID(requestedIID, &IID_IUnknown)){
        *ppv = pMyObj;
        return S_OK;
    }
    return E_NOINTERFACE;
}


static IFileDialogEventsVtbl IFileDialogEventsVtbl_impl = {
    .AddRef = &IFileDialogEvents_AddRef,
    .Release = &IFileDialogEvents_Release,
    .QueryInterface = &IFileDialogEvents_QueryInterface,

    .OnFileOk = &IFileDialog_OnFileOk,
    .OnFolderChange = &IFileDialog_OnFolderChange,
    .OnFolderChanging = &IFileDialog_OnFolderChanging,
    .OnOverwrite = &IFileDialog_OnOverwrite,
    .OnSelectionChange = &IFileDialog_OnSelectionChange,
    .OnShareViolation = &IFileDialog_OnShareViolation,
    .OnTypeChange = &IFileDialog_OnTypeChange
};

static bool work_with_dialog (HWND ownerWnd, const IFileOpenDialog* pDialog, const u16* title
                              , u16 params[MAX_UNICODE_PATH_LENGTH], u16 pathToOpener[MAX_UNICODE_PATH_LENGTH], bool* pSaveConfiguration){
    HRESULT hr;
    {
        // Set the dialog as a folder picker.
        DWORD dwOptions;
        hr = COM_CALL(pDialog, GetOptions, &dwOptions);
        if (! SUCCEEDED(hr)) return false;

        hr = COM_CALL(pDialog, SetOptions, dwOptions | FOS_FORCEFILESYSTEM | FOS_NODEREFERENCELINKS | FOS_DONTADDTORECENT);
        if (! SUCCEEDED(hr)) return false;
    }

    //Set up the dialog

    hr = COM_CALL(pDialog, SetTitle, title);
    if (! SUCCEEDED(hr)) return false;

    static const COMDLG_FILTERSPEC fileTypes[] = {
        {
            .pszName = L"Executable files(*.exe, *.cmd and *.bat)",
            .pszSpec = L"*.exe;*.cmd;*.bat"
        },
        {
            .pszName = L"All Files(*.*)",
            .pszSpec = L"*.*"
        }
    };

    hr = COM_CALL(pDialog, SetFileTypes, sizeof(fileTypes) / sizeof(fileTypes[0]), fileTypes);
    if (! SUCCEEDED(hr)) return false;


    //Customize the dialog

    const IFileDialogCustomize* pCustomizer = customize_dialog(pDialog, params, pSaveConfiguration);
    if (pCustomizer == NULL) return false;

    MyIFileDialogEvents myEventHandler = {
        .lpVtbl = &IFileDialogEventsVtbl_impl,
        .pCustomizer = pCustomizer,
        .params = params,
        .pSaveConfiguration = pSaveConfiguration
    };

    DWORD cookie;
    hr = COM_CALL(pDialog, Advise, (IFileDialogEvents*)&myEventHandler, &cookie);

    //Show the dialog
    hr = COM_CALL(pDialog, Show, ownerWnd);
    if (! SUCCEEDED(hr)){
        COM_CALL0(pCustomizer, Release);
        return false;
    }

    hr = COM_CALL(pDialog, Unadvise, cookie);

    //we no longer need customizer, so
    COM_CALL0(pCustomizer, Release);


    //get path to the handler
    IShellItem* psiResult = NULL;
    hr = COM_CALL(pDialog, GetResult, &psiResult);

    if (SUCCEEDED(hr)){
        u16* pszPath = NULL;
        hr = COM_CALL(psiResult, GetDisplayName, SIGDN_FILESYSPATH, &pszPath);
        COM_CALL0(psiResult, Release);

        if (SUCCEEDED(hr)){
            lstrcpyW(pathToOpener, pszPath);
            CoTaskMemFree(pszPath);
            return true;
        } else {
            return false;
        }
    }

    return true;
}



bool custom_open_with(HWND ownerWnd,const u16* title, u16 params[MAX_UNICODE_PATH_LENGTH], u16 pathToOpener[MAX_UNICODE_PATH_LENGTH], bool* pSaveConfiguration){
    if ( S_OK != CoInitializeEx(NULL, COINIT_MULTITHREADED)) return false;

    bool succeeded = false;

    HRESULT hr = S_OK;
    // Create a new common open file dialog.
    IFileOpenDialog* pfd = NULL;
    hr = CoCreateInstance(&CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, &IID_IFileOpenDialog, (void**)&pfd);
    if (! SUCCEEDED(hr)) return false;
    succeeded = work_with_dialog(ownerWnd, pfd, title, params, pathToOpener, pSaveConfiguration);
    COM_CALL0(pfd, Release);

    return succeeded;
}
