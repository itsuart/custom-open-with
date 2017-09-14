#pragma once
static const GUID SERVER_GUID =
{ 0x2d4f433b, 0xbbe7, 0x48b0, { 0xa9, 0x62, 0xb7, 0x1e, 0x92, 0xe2, 0x2, 0xa4 } };

#define SERVER_GUID_TEXT  L"{2D4F433B-BBE7-48B0-A962-B71E92E202A4}"
#define MAX_UNICODE_PATH_LENGTH 32767

#define COM_CALL(x, y, ...) x->lpVtbl->y(x, __VA_ARGS__)
#define COM_CALL0(x, y) x->lpVtbl->y(x)

typedef uintmax_t uint;
typedef intmax_t sint;
typedef USHORT u16;
typedef unsigned char u8;
