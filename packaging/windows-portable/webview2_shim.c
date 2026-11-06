/*
 * Lets a MinGW build of Folio ship as one .exe.
 *
 * MinGW can't link Microsoft's static WebView2 loader, so it would normally
 * need WebView2Loader.dll next to the program. This shim carries that DLL
 * inside the .exe, unpacks it once to the user's local app data, and forwards
 * the loader's five exported functions to it. The DLL is Microsoft's, unmodified.
 */
#include <windows.h>

#define WSTR2(x) L##x
#define WSTR(x) WSTR2(x)

__asm__(".section .rdata,\"dr\"\n"
        ".globl folio_wv2_start\nfolio_wv2_start:\n"
        ".incbin \"" LOADER_PATH "\"\n"
        ".globl folio_wv2_end\nfolio_wv2_end:\n"
        ".text\n");
extern const unsigned char folio_wv2_start[], folio_wv2_end[];

static BOOL file_matches(const wchar_t *path, const unsigned char *data, DWORD size) {
    HANDLE f = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, 0, NULL);
    if (f == INVALID_HANDLE_VALUE) return FALSE;
    BOOL same = GetFileSize(f, NULL) == size;
    unsigned char buf[8192];
    DWORD at = 0, got = 0;
    while (same && at < size && ReadFile(f, buf, sizeof buf, &got, NULL) && got) {
        for (DWORD i = 0; i < got; i++) if (buf[i] != data[at + i]) { same = FALSE; break; }
        at += got;
    }
    CloseHandle(f);
    return same && at == size;
}

static BOOL write_file(const wchar_t *path, const unsigned char *data, DWORD size) {
    HANDLE f = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (f == INVALID_HANDLE_VALUE) return FALSE;
    DWORD put = 0;
    BOOL ok = WriteFile(f, data, size, &put, NULL) && put == size;
    CloseHandle(f);
    return ok;
}

static HMODULE unpack_and_load(void) {
    const DWORD size = (DWORD)(folio_wv2_end - folio_wv2_start);
    wchar_t dir[MAX_PATH], path[MAX_PATH + 64], tmp[MAX_PATH + 64];
    DWORD n = GetEnvironmentVariableW(L"LOCALAPPDATA", dir, MAX_PATH);
    if (!n || n >= MAX_PATH) return NULL;
    lstrcatW(dir, L"\\com.kalystron.folio");
    CreateDirectoryW(dir, NULL);
    lstrcpyW(path, dir);
    lstrcatW(path, L"\\WebView2Loader.dll");
    if (!file_matches(path, folio_wv2_start, size)) {
        lstrcpyW(tmp, path);
        lstrcatW(tmp, L".new");
        if (!write_file(tmp, folio_wv2_start, size)) return NULL;
        if (!MoveFileExW(tmp, path, MOVEFILE_REPLACE_EXISTING)) { DeleteFileW(tmp); return NULL; }
    }
    return LoadLibraryExW(path, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
}

static FARPROC resolve(const char *name) {
    static HMODULE mod;
    if (!mod) mod = unpack_and_load();
    return mod ? GetProcAddress(mod, name) : NULL;
}

#define MISSING HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND)
#define FORWARD(name, params, args)                                   \
    typedef HRESULT(WINAPI *name##_t) params;                         \
    HRESULT WINAPI name params {                                      \
        name##_t fn = (name##_t)(void *)resolve(#name);               \
        return fn ? fn args : MISSING;                                \
    }                                                                 \
    void *__imp_##name = (void *)name;

FORWARD(CompareBrowserVersions, (LPCWSTR a, LPCWSTR b, int *r), (a, b, r))
FORWARD(CreateCoreWebView2Environment, (void *h), (h))
FORWARD(CreateCoreWebView2EnvironmentWithOptions, (LPCWSTR a, LPCWSTR b, void *c, void *d), (a, b, c, d))
FORWARD(GetAvailableCoreWebView2BrowserVersionString, (LPCWSTR a, LPWSTR *b), (a, b))
FORWARD(GetAvailableCoreWebView2BrowserVersionStringWithOptions, (LPCWSTR a, void *b, LPWSTR *c), (a, b, c))
