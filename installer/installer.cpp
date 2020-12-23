#define _WIN32_WINNT _WIN32_WINNT_WIN7
#include <Windows.h>
#include <Shlwapi.h>
#include <guiddef.h>
#include <Shlobj.h>

#include <cstdint>
#include <string>
#include <memory>

namespace {

    const wchar_t* EXTENSION_GUID_TEXT{ L"{7BA11196-950C-4CC8-81E8-9853F514127F}" };

    const wchar_t* EXTENSION_CLSID_TEXT{ L"CLSID\\{7BA11196-950C-4CC8-81E8-9853F514127F}" };

    const wchar_t* EXTENSION_REGKEY_TEXT{
        L"AllFileSystemObjects\\ShellEx\\ContextMenuHandlers\\{7BA11196-950C-4CC8-81E8-9853F514127F}"
    };

    const wchar_t* EXTENSION_OPEN_BG_DIRECTORY_REGKEY_TEXT{
        L"Directory\\Background\\ShellEx\\ContextMenuHandlers\\{7BA11196-950C-4CC8-81E8-9853F514127F}"
    };

    constexpr GUID EXTENSION_GUID = { 0x7ba11196, 0x950c, 0x4cc8,{ 0x81, 0xe8, 0x98, 0x53, 0xf5, 0x14, 0x12, 0x7f } };

    const wchar_t* DLL_NAME{ L"extension.dll" };

    using uint = uintmax_t;

    constexpr int MAX_WIDE_PATH_LENGTH = 32767;
    const std::wstring UNC_PATH_PREFIX(LR"(\\)");
    const std::wstring WIDE_PATH_PREFIX(LR"(\\?\)");

    const wchar_t* PROGNAME{ L"'My Open With' Shell Extension (Un)Installer" };
    void display_error(DWORD error) {
        wchar_t* reason = NULL;

        FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, 0, (LPWSTR)&reason, 1, NULL);
        MessageBoxW(NULL, reason, PROGNAME, MB_OK | MB_ICONERROR);
        HeapFree(GetProcessHeap(), 0, reason);
    }

    void display_last_error() {
        DWORD last_error = GetLastError();
        return display_error(last_error);
    }

    void inform(const wchar_t* message) {
        MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONINFORMATION);
    }

    bool ask(const wchar_t* question) {
        return IDYES == MessageBoxW(NULL, question, PROGNAME, MB_YESNO | MB_ICONQUESTION);
    }

    void error(const wchar_t* message) {
        MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONERROR);
    }

    void warn(const wchar_t* message) {
        MessageBoxW(NULL, message, PROGNAME, MB_OK | MB_ICONWARNING);
    }

    bool register_com_server(const std::wstring& server_path) {
        static const wchar_t description[] = L"'My Open With' Shell Extension";

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

        errorCode = RegSetValueExW(inprocKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"both", sizeof(L"both"));
        if (errorCode != ERROR_SUCCESS) {
            display_error(errorCode);
            return false;
        }


        RegCloseKey(inprocKey);
        RegCloseKey(clsidKey);
        return true;
    }

    bool unregister_com_server(bool silent) {
        long errorCode = RegDeleteTreeW(HKEY_CLASSES_ROOT, EXTENSION_CLSID_TEXT);
        if (errorCode != ERROR_SUCCESS) {
            if (!silent) {
                display_error(errorCode);
            }
            return false;
        }
        return true;
    }

    bool register_shell_extension() {
        {
            HKEY key = nullptr;
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
            if (errorCode != ERROR_SUCCESS) {
                display_error(errorCode);
                return false;
            }
            ::RegCloseKey(key);
        }

        {
            HKEY key = nullptr;
            LONG errorCode = RegCreateKeyExW(
                                             HKEY_CLASSES_ROOT,
                                             EXTENSION_OPEN_BG_DIRECTORY_REGKEY_TEXT,
                                             0, NULL,
                                             REG_OPTION_NON_VOLATILE,
                                             KEY_ALL_ACCESS,
                                             NULL,
                                             &key,
                                             NULL
                                             );
            if (errorCode != ERROR_SUCCESS) {
                display_error(errorCode);
                ::RegDeleteTreeW(HKEY_CLASSES_ROOT, EXTENSION_REGKEY_TEXT);
                return false;
            }
            ::RegCloseKey(key);
        }

        return true;
    };

    bool unregister_shell_extension(bool silent) {
        bool hadError = false;
        for (const wchar_t* key : { EXTENSION_REGKEY_TEXT, EXTENSION_OPEN_BG_DIRECTORY_REGKEY_TEXT }) {
            long errorCode = RegDeleteTreeW(HKEY_CLASSES_ROOT, EXTENSION_REGKEY_TEXT);
            if (errorCode != ERROR_SUCCESS) {
                hadError |= true;
                if (!silent) {
                    display_error(errorCode);
                }
            }

        }
        return (not hadError);
    }

    std::unique_ptr<wchar_t, decltype(CoTaskMemFree)*> GetUserDocumentsFolderPath() {
        wchar_t* pMyDocuments = nullptr;
        SHGetKnownFolderPath(FOLDERID_Documents, KF_FLAG_CREATE, 0, &pMyDocuments);
        return std::unique_ptr<wchar_t, decltype(CoTaskMemFree)*>(pMyDocuments, CoTaskMemFree);
    }

    bool delete_folders_for_handlers(bool silent) {
        auto myDocuments = GetUserDocumentsFolderPath();
        std::wstring workingString(myDocuments.get());

        /* yeah, but MSDN says, that :"SHFileOperation fails on any path prefixed with "\\?\"." and it correct, so..
        if (workingString.substr(0, UNC_PATH_PREFIX.length()) != UNC_PATH_PREFIX) {
            workingString.insert(0, WIDE_PATH_PREFIX);
        }
        */
        using namespace std::string_literals; // gotta love c++14 :)
        workingString.append(L"\\Open With Handlers for\0"s);

        SHFILEOPSTRUCTW op = { 0 };
        op.hwnd = NULL;
        op.wFunc = FO_DELETE;
        op.pFrom = workingString.c_str();
        op.pTo = NULL;
        op.fFlags = silent ? FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_NOCONFIRMMKDIR | FOF_ALLOWUNDO : FOF_ALLOWUNDO;
        auto ret = SHFileOperationW(&op);
        return 0 == ret;
    }

    bool uninstall_extension(bool silent) {
        bool allSucceeded = true;
        allSucceeded &= unregister_com_server(silent);
        allSucceeded &= unregister_shell_extension(silent);
        if (silent || ask(L"Move your handlers folder into Trash (Recycle Bin)?")) {
            allSucceeded &= delete_folders_for_handlers(silent);
        }

        return allSucceeded;
    }

    bool create_folder_if_not_exists(const std::wstring& path) {
        if (!CreateDirectoryW(path.c_str(), NULL)) {
            return ERROR_ALREADY_EXISTS == GetLastError();
        }
        return true;
    }

    bool create_folders_for_handlers(std::wstring& pathToHandlersFolder) {
        auto myDocuments = GetUserDocumentsFolderPath();
        std::wstring workingString(myDocuments.get());

        if (workingString.substr(0, UNC_PATH_PREFIX.length()) != UNC_PATH_PREFIX) {
            workingString.insert(0, WIDE_PATH_PREFIX);
        }

        workingString.append(L"\\Open With Handlers for");
        pathToHandlersFolder = workingString;
        const auto rootLength = workingString.length();

        if (!create_folder_if_not_exists(workingString)) return false;

        for (const wchar_t* folderName : { L"\\Everything", L"\\Folders", L"\\All files", L"\\Extensionless Files", L"\\Files by Extension" }) {
            workingString.append(folderName);
            if (!create_folder_if_not_exists(workingString)) return false;
            workingString.erase(rootLength);
        }

        return true;
    }

    bool install_extension(const std::wstring& server_path, std::wstring& pathToHandlersFolder) {
        if (!create_folders_for_handlers(pathToHandlersFolder)) {
            return false;
        }

        if (!register_com_server(server_path)) {
            return false;
        }

        if (!register_shell_extension()) {
            return false;
        }

        return true;
    }

    bool is_extension_installed() {
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
                    uint flag = 0xFFLL << (i * 8);
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

    // erases everything after LAST backslash (\)
    // if there is no backslashes - does nothing
    std::wstring& chop_off_filename(std::wstring& path) {
        const auto pos = path.rfind('\\');
        if (pos != std::wstring::npos && (pos + 1 < path.length())) {
            path.erase(pos + 1);
        }
        return path;
    }

}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, wchar_t* lpCmdLine, int nCmdShow){
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(nCmdShow);
    UNREFERENCED_PARAMETER(lpCmdLine);

    if (CoInitializeEx(NULL, COINIT_MULTITHREADED) != S_OK) {
        error(L"Failed to initialize COM");
        return 1;
    }

    const wchar_t* commandLine = GetCommandLineW();
    int nArgs = 0;
    const wchar_t** args = const_cast<const wchar_t**>(CommandLineToArgvW(commandLine, &nArgs));
    if (args == NULL) {
        display_last_error();
        return 1;
    }

    std::wstring processFullPath;
    processFullPath.resize(MAX_WIDE_PATH_LENGTH);
    DWORD processFullPathLength = GetModuleFileNameW(NULL, &processFullPath.front(), processFullPath.size());
    processFullPath.resize(processFullPathLength);

    if (nArgs == 1) {
        if (is_extension_installed()) {
            if (ask(L"The extension is installed, would you like to UNINSTALL it?")) {
                ShellExecuteW(NULL, L"runas", processFullPath.data(), L"u", NULL, SW_SHOWDEFAULT);
            }
        } else {
            if (ask(L"Would you like to INSTALL the extension?")) {
                ShellExecuteW(NULL, L"runas", processFullPath.data(), L"i", NULL, SW_SHOWDEFAULT);
            }
        }
        return 0;
    } else {
        //one or more args were passed. take first arg and see what is it
        const wchar_t* arg = args[1];
        switch (*arg) {
            case L'i': {
                chop_off_filename(processFullPath);
                processFullPath.append(DLL_NAME);

                std::wstring pathToHandlersFolder;
                if (install_extension(processFullPath.c_str(), pathToHandlersFolder)) {
                    if (ask(L"The extension was registered successfuly.\nWould you like to open handlers folders?")) {
                        ShellExecuteW(NULL, L"explore", pathToHandlersFolder.c_str(), NULL, NULL, SW_SHOWDEFAULT);
                    }
                } else {
                    uninstall_extension(true);
                };
            } break;

            case L'u': {
                if (uninstall_extension(false)) {
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
