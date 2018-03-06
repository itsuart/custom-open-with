#define _WIN32_WINNT _WIN32_WINNT_WIN7
#include <Windows.h>
#include <Shlwapi.h>
#include <guiddef.h>

#include <cstdint>
#include <array>
#include <string>

static const wchar_t* EXTENSION_GUID_TEXT{ L"{7BA11196-950C-4CC8-81E8-9853F514127F}" };
static const wchar_t* EXTENSION_CLSID_TEXT{ L"CLSID\\{7BA11196-950C-4CC8-81E8-9853F514127F}" };
static const wchar_t* EXTENSION_REGKEY_TEXT{ L"AllFileSystemObjects\\ShellEx\\ContextMenuHandlers\\{7BA11196-950C-4CC8-81E8-9853F514127F}" };

static constexpr GUID EXTENSION_GUID = { 0x7ba11196, 0x950c, 0x4cc8,{ 0x81, 0xe8, 0x98, 0x53, 0xf5, 0x14, 0x12, 0x7f } };

static const wchar_t DLL_NAME[] = { L"extension.dll" };

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;
typedef unsigned char u8;


static void display_error(DWORD error) {
    wchar_t* reason = NULL;

    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, error, 0, (LPWSTR)&reason, 1, NULL);
    MessageBoxW(NULL, reason, NULL, MB_OK);
    HeapFree(GetProcessHeap(), 0, reason);
}

static void display_last_error() {
    DWORD last_error = GetLastError();
    return display_error(last_error);
}

// erases everything after LAST backslash (\)
// if there is no backslashes - does nothing
static std::wstring& chop_off_filename(std::wstring& path) {
    const auto pos = path.rfind('\\');
    if (pos != std::wstring::npos && (pos + 1 < path.length())) {
        path.erase(pos + 1);
    }
    return path;
}

static const wchar_t* PROGNAME{ L"'My Open With...' Shell Extension (Un)Installer" };
static void inform(const wchar_t* message) {
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONINFORMATION);
}

static bool ask(const wchar_t* question) {
    return IDYES == MessageBoxW(NULL, question, PROGNAME, MB_YESNO | MB_ICONQUESTION);
}

static void error(const wchar_t* message) {
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONERROR);
}

static void warn(const wchar_t* message) {
    MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONWARNING);
}

static bool register_com_server(const std::wstring& server_path) {
    static const wchar_t description[] = L"'My Open With...' Shell Extension";

    HKEY clsidKey;
    LONG errorCode = RegCreateKeyExW(
        HKEY_CLASSES_ROOT,
        EXTENSION_CLSID_TEXT,
        0, NULL,
        REG_OPTION_NON_VOLATILE,
        KEY_ALL_ACCESS,
        NULL,
        &clsidKey,
        NULL
    );
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }
    errorCode = RegSetValueExW(clsidKey, NULL, 0, REG_SZ, (BYTE*)description, sizeof(description));
    if (errorCode != ERROR_SUCCESS) {
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

    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }

    errorCode = RegSetValueExW(inprocKey, NULL, 0, REG_SZ, reinterpret_cast<const BYTE*>(server_path.c_str()), (server_path.length() + 1) * sizeof(wchar_t));
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }

    errorCode = RegSetValueExW(inprocKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"both", sizeof(L"both") + 2);
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }


    RegCloseKey(inprocKey);
    RegCloseKey(clsidKey);
    return true;
}

static bool register_shell_extension() {
    HKEY key;
    LONG errorCode = RegCreateKeyExW(
        HKEY_CLASSES_ROOT,
        EXTENSION_REGKEY_TEXT,
        0, NULL,
        REG_OPTION_NON_VOLATILE,
        KEY_ALL_ACCESS,
        NULL,
        &key,
        NULL
    );
    RegCloseKey(key);
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }
    return true;
};

static bool unregister_shell_extension() {
    long errorCode = RegDeleteTreeW( HKEY_CLASSES_ROOT, EXTENSION_REGKEY_TEXT);
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }
    return true;
}

static bool uninstall_extension() {
    long errorCode = RegDeleteTreeW(HKEY_CLASSES_ROOT, EXTENSION_CLSID_TEXT);
    if (errorCode != ERROR_SUCCESS) {
        display_error(errorCode);
        return false;
    }
    return unregister_shell_extension();
}

static bool install_extension(const std::wstring& server_path) {
    if (register_com_server(server_path)) {
        return register_shell_extension();
    };
    return false;
}

static bool is_extension_installed() {
    if (CoInitializeEx(NULL, COINIT_MULTITHREADED) != S_OK) {
        error(L"Failed to initialize COM");
        return false;
    }

    IUnknown* pIUnknown = NULL;
    HRESULT result = CoCreateInstance(EXTENSION_GUID, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&pIUnknown);
    switch (result) {
        case REGDB_E_CLASSNOTREG: {
            return false;
        } break;

        case CLASS_E_NOAGGREGATION: {
            error(L"CLASS_E_NOAGGREGATION");
        } break;

        case E_NOINTERFACE: {
            error(L"E_NOINTERFACE");
        } break;

        case E_POINTER: {
            error(L"E_POINTER");
        } break;

        case E_OUTOFMEMORY: {
            error(L"E_OUTOFMEMORY");
        } break;

        case S_OK: {
            pIUnknown->Release();
            return true;
        } break;

        default: {
            const static wchar_t alphabet[] = L"0123456789ABCDEF";
            wchar_t errorCode[2 * sizeof(HRESULT) + 10] = { 0 };
            for (int i = sizeof(HRESULT); i >= 0; i -= 1) {
                uint flag = 0xFF << (i * 8);
                uint part = (result & flag) >> (i * 8);
                uint high = (part & 0xF0) / 0x10;
                uint low = (part & 0x0F);
                errorCode[2 * (sizeof(HRESULT) - i)] = alphabet[high];
                errorCode[2 * (sizeof(HRESULT) - i) + 1] = alphabet[low];
            }
            warn(errorCode);
        }
    }

    return false;
}


int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, wchar_t* lpCmdLine, int nCmdShow){
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(nCmdShow);
    UNREFERENCED_PARAMETER(lpCmdLine);

    
    static std::array<wchar_t, MAX_PATH> buffer{0};
    DWORD processFullPathLength = GetModuleFileNameW(NULL, buffer.data(), buffer.size());
    static std::wstring processFullPath(buffer.data());

    const wchar_t* commandLine = GetCommandLineW();
    int nArgs = 0;
    const wchar_t** args = const_cast<const wchar_t**>(CommandLineToArgvW(commandLine, &nArgs));
    if (args == NULL) {
        display_last_error();
        return 1;
    }

    if (nArgs == 1) {
        if (is_extension_installed()) {
            if (ask(L"The extension is installed, would you like to UNINSTALL it?")) {
                ShellExecuteW(NULL, L"runas", processFullPath.c_str(), L"u", NULL, SW_SHOWDEFAULT);
            }
        } else {
            if (ask(L"Would you like to INSTALL the extension?")) {
                ShellExecuteW(NULL, L"runas", processFullPath.c_str(), L"i", NULL, SW_SHOWDEFAULT);
            }
        }

    } else {
        //one or more args were passed. take first arg and see what is it
        const wchar_t* arg = args[1];
        switch (*arg) {
            case L'i': {
                chop_off_filename(processFullPath);
                processFullPath.append(DLL_NAME);

                if (install_extension(processFullPath.c_str())) {
                    inform(L"The extension was registered successfuly.");
                };
            } break;

            case L'u': {
                if (uninstall_extension()) {
                    inform(L"The extension was unregistered, but some applications might still use it. You should be able to remove files after reboot.");
                };
            } break;

            default: {
                error(L"Invalid command line");
            }
        }
    }

    return 0;
}
