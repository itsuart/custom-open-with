#pragma once
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

#include "common.c"
/*
typedef struct {
    void* pSomething;
} ShellLink;

static const uint MAX_SHELLLINK_ARGS_LENGTH = 32 * 1024;

typedef struct tag_ShellLink{
    u16 fullPath[MAX_PATH];
    u16 args[MAX_SHELLLINK_ARGS_LENGTH];
    а16 description[INFOTIPSIZE];
    WORD hotkey;

} ShellLink;

bool Write_ShellLink(ShellLink link, u16* targetFile);

bool Read_ShellLink(u16* pathToShellLinkFile, ShellLink* pResult);


bool ShellLink_Load(u16* pathToLinkFile, ShellLink* pResult);

*/
bool ShellLink_CreateAndStore(u16* linkTargetPath, u16* linkTargetArguments, u16* whereToStore);
