#include <Windows.h>

const wchar_t* __stdcall exported_function(const wchar_t* input) {
    return input;
}

BOOL APIENTRY entry_point(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
    UNREFERENCED_PARAMETER(hModule);
    UNREFERENCED_PARAMETER(lpReserved);
    switch (ul_reason_for_call) {
        case DLL_PROCESS_ATTACH: {
        } break;

        case DLL_THREAD_ATTACH: {
        } break;

        case DLL_THREAD_DETACH: {
        } break;

        case DLL_PROCESS_DETACH: {

        } break;

        default: {

        } break;

    }
    return TRUE;
}
