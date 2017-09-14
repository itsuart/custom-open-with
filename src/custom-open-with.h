#pragma once
#include "common.c"
#include <windows.h>

bool custom_open_with(HWND ownerWnd,const u16* title, u16 params[MAX_UNICODE_PATH_LENGTH], u16 pathToOpener[MAX_UNICODE_PATH_LENGTH], bool* pSaveConfiguration);
