#define UNICODE
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#include <windows.h>
#include <tchar.h>
#include <stdbool.h>
#include <stdint.h>
#include <GuidDef.h>
#include <winerror.h>

/* GLOBALS */
#include "common.c"

#include "custom-open-with.h"
#include "shell-link.h"

static void display_error(DWORD error){
    u16* reason = NULL;

    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, error, 0, (LPWSTR)&reason, 1, NULL);
    MessageBoxW(NULL, reason, NULL, MB_OK);
    HeapFree(GetProcessHeap(), 0, reason);
}

static void display_last_error(){
    DWORD last_error = GetLastError();
    return display_error(last_error);
}

static uint remove_last_part_from_path(u16* full_path, uint full_path_length){
    while (full_path_length >= 0){
        u16* current_ptr = full_path + full_path_length;
        u16 current_char = *current_ptr;
        if (current_char == L'\\'){
            return full_path_length + 1;
        } else {
            *current_ptr = 0;
            full_path_length -= 1;
        }
    }
}

#define PROGNAME L"Custom Open With Shell Extension In/Unstaller"

static const u16 extension_name[] = L"extension.dll\0";

static void inform (u16* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONINFORMATION);
}

static bool ask(u16* question){
    return IDYES == MessageBoxW(NULL, question, PROGNAME, MB_YESNO | MB_ICONQUESTION);
}

static void error(u16* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONERROR);
}

static void warn(u16* message){
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONWARNING);
}

static bool register_com_server(u16* server_path, uint server_path_length){
    static u16 description[] = L"Windows Explorer context menu extension to copy, move and duplicate files and directories\0";

    HKEY clsidKey;
    LONG errorCode = RegCreateKeyExW(
         HKEY_CLASSES_ROOT,
         L"CLSID\\" SERVER_GUID_TEXT,
         0, NULL,
         REG_OPTION_NON_VOLATILE,
         KEY_ALL_ACCESS,
         NULL,
         &clsidKey,
         NULL
    );
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }
    errorCode = RegSetValueExW(clsidKey, NULL, 0, REG_SZ, (BYTE*)description, sizeof(description));
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }

    HKEY inprocKey;
    errorCode = RegCreateKeyExW(
              clsidKey,
              L"InprocServer32",
              0, NULL,
              REG_OPTION_NON_VOLATILE,
              KEY_ALL_ACCESS,
              NULL,
              &inprocKey,
              NULL
    );

    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }

    errorCode = RegSetValueExW(inprocKey, NULL, 0, REG_SZ, (BYTE*)server_path, (server_path_length + 1) * sizeof(u16));
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }

    errorCode = RegSetValueExW(inprocKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"both", sizeof(L"both") + 2);
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }


    RegCloseKey(inprocKey);
    RegCloseKey(clsidKey);
    return true;
}

static bool register_shell_extension(){
    HKEY key;
    LONG errorCode = RegCreateKeyExW(
         HKEY_CLASSES_ROOT,
         L"AllFileSystemObjects\\ShellEx\\ContextMenuHandlers\\" SERVER_GUID_TEXT,
         0, NULL,
         REG_OPTION_NON_VOLATILE,
         KEY_ALL_ACCESS,
         NULL,
         &key,
         NULL
    );
    RegCloseKey(key);
    if (errorCode != ERROR_SUCCESS){
       display_error(errorCode);
       return false;
    }
    return true;
};

static bool unregister_shell_extension(){
    long errorCode = RegDeleteTreeW(
           HKEY_CLASSES_ROOT,
           L"AllFileSystemObjects\\ShellEx\\ContextMenuHandlers\\" SERVER_GUID_TEXT);
    if (errorCode != ERROR_SUCCESS){
        display_error(errorCode);
        return false;
    }
    return true;
}

static bool uninstall_extension(){
    long errorCode = RegDeleteTreeW(HKEY_CLASSES_ROOT, L"CLSID\\" SERVER_GUID_TEXT);
    if (errorCode != ERROR_SUCCESS){
        display_error(errorCode);
        return false;
    }
    return unregister_shell_extension();
}

static bool install_extension(u16* server_path){
    if (register_com_server(server_path, lstrlenW(server_path))){
        return register_shell_extension();
    };
    return false;
}

static bool is_extension_installed(){
    HRESULT result = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    IUnknown* pIUnknown = NULL;

    if (result != S_OK){
        error(L"Failed to initialize COM");
        return false;
    }
    result = CoCreateInstance(&SERVER_GUID, NULL, CLSCTX_INPROC_SERVER, &IID_IUnknown, (void**)&pIUnknown);
    switch (result){
        case REGDB_E_CLASSNOTREG:{
            return false;
        } break;

        case CLASS_E_NOAGGREGATION:{
            error(L"CLASS_E_NOAGGREGATION");
        } break;

        case E_NOINTERFACE:{
            error(L"E_NOINTERFACE");
        } break;

        case E_POINTER:{
            error(L"E_POINTER");
        } break;

        case E_OUTOFMEMORY:{
            error(L"E_OUTOFMEMORY");
        } break;

        case S_OK:{
            pIUnknown->lpVtbl->Release(pIUnknown);
            return true;
        } break;

        default:{
            u16 alphabet[16] = L"0123456789ABCDEF";
            u16 errorCode[2*sizeof(HRESULT) + 10] = {0};
            for (int i = sizeof(HRESULT); i >= 0 ; i-= 1){
                uint flag = 0xFF << (i*8);
                uint part = (result & flag) >> (i*8);
                uint high = (part & 0xF0) / 0x10;
                uint low = (part & 0x0F);
                errorCode[2*(sizeof(HRESULT) - i)] = alphabet[high];
                errorCode[2*(sizeof(HRESULT) - i) + 1] = alphabet[low];
            }
            warn(errorCode);
        }
    }

    return false;
}

static bool saveConfig = false;
static u16 params[MAX_UNICODE_PATH_LENGTH] = {0};
static u16 pathToOpener[MAX_UNICODE_PATH_LENGTH] = {0};

static u16 processFullPath[MAX_UNICODE_PATH_LENGTH + 1] = {0};
static uint processFullPathLength = 0;
void entry_point(){
    custom_open_with(NULL, L"Title", params, pathToOpener, &saveConfig);
    if (saveConfig){
        ShellLink_CreateAndStore(pathToOpener, params, L"c:\\tmp\\baz.lnk");
        ShellLink_CreateAndStore(pathToOpener, params, L"c:\\tmp\\foo\\bar\\baz.lnk");
    }
    ExitProcess(0);
    return;

    processFullPathLength = GetModuleFileNameW(NULL, processFullPath, MAX_UNICODE_PATH_LENGTH);
    if (processFullPathLength == 0){
        display_last_error();
        ExitProcess(1);
        return;
    }

    const u16* commandLine = GetCommandLineW();

    int nArgs = 0;
    u16** args = CommandLineToArgvW(commandLine, &nArgs);
    if (args == NULL){
        display_last_error();
        ExitProcess(1);
        return;
    }

    if (nArgs == 1){
        if (is_extension_installed()){
            if (ask(L"Shell extension is installed, would you like to UNINSTALL it?")){
                ShellExecuteW(NULL, L"runas", processFullPath, L"u", NULL, SW_SHOWDEFAULT);
            }
        } else {
            if (ask(L"Would you like to INSTALL the extension?")){
                ShellExecuteW(NULL, L"runas", processFullPath, L"i", NULL, SW_SHOWDEFAULT);
            }
        }

    } else {
        //one or more args were passed. take first arg and see what is it
        u16* arg = args[1];
        switch(*arg){
           case L'i':{
               uint length = remove_last_part_from_path(processFullPath, processFullPathLength);
               CopyMemory(processFullPath + length, extension_name, sizeof(extension_name));

               if (install_extension(processFullPath)){
                   inform(L"Extension was registered successfuly.");
               };
           } break;

           case L'u':{
               if (uninstall_extension()){
                   inform(L"The extension was unregistered, but some applications might still use it. You should be able to remove files after reboot.");
               };
           } break;

           default:{
               error(L"Invalid command line");
           }
        }
    }

    ExitProcess(0);
    return;
}
