#include <Windows.h>

static const wchar_t* EXTENSION_GUID_TEXT{ L"{7BA11196-950C-4CC8-81E8-9853F514127F}" };
static constexpr GUID EXTENSION_GUID =
{ 0x7ba11196, 0x950c, 0x4cc8,{ 0x81, 0xe8, 0x98, 0x53, 0xf5, 0x14, 0x12, 0x7f } };

const wchar_t* __stdcall exported_function(const wchar_t* input) {
    return input;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

