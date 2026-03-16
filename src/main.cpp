#define UNICODE
#define _UNICODE

#include <windows.h>
#include <shellapi.h>
#include <strsafe.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>
#include <wctype.h>

typedef struct registry_location_t {
    const wchar_t *scope_name;
    const wchar_t *root_name;
    HKEY root_key;
    const wchar_t *subkey;
    const wchar_t *entry_type;
} REGISTRY_LOCATION;

typedef struct scope_t {
    const wchar_t *name;
    const wchar_t *description;
    const REGISTRY_LOCATION *locations;
    int location_count;
} SCOPE;

typedef int (*ENTRY_CALLBACK)(const REGISTRY_LOCATION *location, const wchar_t *entry_name, const wchar_t *display, const wchar_t *command, const wchar_t *path, void *context);

typedef struct export_context_t {
    FILE *file;
    int write_failed;
    const struct preserved_export_state_t *preserved_state;
    int written_k_entries;
    int skipped_preserved_entries;
} EXPORT_CONTEXT;

typedef struct preserved_export_entry_t {
    wchar_t *line;
    wchar_t *path;
} PRESERVED_EXPORT_ENTRY;

typedef struct preserved_export_state_t {
    PRESERVED_EXPORT_ENTRY *entries;
    int count;
    int capacity;
    int malformed_count;
} PRESERVED_EXPORT_STATE;

typedef enum builtin_action_kind_t {
    BUILTIN_ACTION_BLOCK_CLSID,
    BUILTIN_ACTION_REMOVE_PATHS
} BUILTIN_ACTION_KIND;

typedef struct builtin_target_t {
    const wchar_t *name;
    const wchar_t *label;
    BUILTIN_ACTION_KIND action_kind;
    const wchar_t *const *values;
    int value_count;
} BUILTIN_TARGET;

static const wchar_t *g_builtin_scope_name = L"builtin";
static const wchar_t *g_builtin_path_prefix = L"builtin:";

static const REGISTRY_LOCATION g_file_locations[] = {
    {L"file", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\*\\shell", L"shell"},
    {L"file", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\*\\shellex\\ContextMenuHandlers", L"handler"},
    {L"file", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\*\\shell", L"shell"},
    {L"file", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\*\\shellex\\ContextMenuHandlers", L"handler"},
};

static const REGISTRY_LOCATION g_folder_locations[] = {
    {L"folder", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Directory\\shell", L"shell"},
    {L"folder", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Directory\\shellex\\ContextMenuHandlers", L"handler"},
    {L"folder", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Folder\\shell", L"shell"},
    {L"folder", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Folder\\shellex\\ContextMenuHandlers", L"handler"},
    {L"folder", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Directory\\shell", L"shell"},
    {L"folder", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Directory\\shellex\\ContextMenuHandlers", L"handler"},
    {L"folder", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Folder\\shell", L"shell"},
    {L"folder", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Folder\\shellex\\ContextMenuHandlers", L"handler"},
};

static const REGISTRY_LOCATION g_desktop_locations[] = {
    {L"desktop", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\DesktopBackground\\shell", L"shell"},
    {L"desktop", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\DesktopBackground\\shellex\\ContextMenuHandlers", L"handler"},
    {L"desktop", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\DesktopBackground\\shell", L"shell"},
    {L"desktop", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\DesktopBackground\\shellex\\ContextMenuHandlers", L"handler"},
};

static const REGISTRY_LOCATION g_explorer_locations[] = {
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Directory\\Background\\shell", L"shell"},
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Directory\\Background\\shellex\\ContextMenuHandlers", L"handler"},
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\AllFilesystemObjects\\shell", L"shell"},
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\AllFilesystemObjects\\shellex\\ContextMenuHandlers", L"handler"},
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Drive\\shell", L"shell"},
    {L"explorer", L"HKCU", HKEY_CURRENT_USER, L"Software\\Classes\\Drive\\shellex\\ContextMenuHandlers", L"handler"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Directory\\Background\\shell", L"shell"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Directory\\Background\\shellex\\ContextMenuHandlers", L"handler"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\AllFilesystemObjects\\shell", L"shell"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\AllFilesystemObjects\\shellex\\ContextMenuHandlers", L"handler"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Drive\\shell", L"shell"},
    {L"explorer", L"HKLM", HKEY_LOCAL_MACHINE, L"Software\\Classes\\Drive\\shellex\\ContextMenuHandlers", L"handler"},
};

static const SCOPE g_scopes[] = {
    {L"file", L"File right-click menu entries", g_file_locations, (int)(sizeof(g_file_locations) / sizeof(g_file_locations[0]))},
    {L"folder", L"Folder right-click menu entries", g_folder_locations, (int)(sizeof(g_folder_locations) / sizeof(g_folder_locations[0]))},
    {L"desktop", L"Desktop background right-click menu entries", g_desktop_locations, (int)(sizeof(g_desktop_locations) / sizeof(g_desktop_locations[0]))},
    {L"explorer", L"Explorer background, drive, and filesystem entries", g_explorer_locations, (int)(sizeof(g_explorer_locations) / sizeof(g_explorer_locations[0]))},
};

static const wchar_t *g_builtin_photos_edit_clsids[] = {
    L"{BFE0E2A4-C70C-4AD7-AC3D-10D1ECEBB5B4}",
};

static const wchar_t *g_builtin_designer_clsids[] = {
    L"{7A53B94A-4E6E-4826-B48E-535020B264E5}",
};

static const wchar_t *g_builtin_copilot_clsids[] = {
    L"{CB3B0003-8088-4EDE-8769-8B354AB2FF8C}",
};

static const wchar_t *g_builtin_defender_clsids[] = {
    L"{09A47860-11B0-4DA5-AFA5-26D86198A780}",
};

static const wchar_t *g_builtin_setdesktopwallpaper_paths[] = {
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.avci\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.avcs\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.avif\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.avifs\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.bmp\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.dib\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.gif\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.heic\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.heics\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.heif\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.heifs\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.jfif\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.jpe\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.jpeg\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.jpg\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.png\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.tif\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.tiff\\Shell\\setdesktopwallpaper",
    L"HKLM\\Software\\Classes\\SystemFileAssociations\\.wdp\\Shell\\setdesktopwallpaper",
};

static const BUILTIN_TARGET g_builtin_targets[] = {
    {L"photos-edit", L"Edit with Photos", BUILTIN_ACTION_BLOCK_CLSID, g_builtin_photos_edit_clsids, (int)(sizeof(g_builtin_photos_edit_clsids) / sizeof(g_builtin_photos_edit_clsids[0]))},
    {L"designer", L"Create with Designer", BUILTIN_ACTION_BLOCK_CLSID, g_builtin_designer_clsids, (int)(sizeof(g_builtin_designer_clsids) / sizeof(g_builtin_designer_clsids[0]))},
    {L"copilot", L"Ask Copilot", BUILTIN_ACTION_BLOCK_CLSID, g_builtin_copilot_clsids, (int)(sizeof(g_builtin_copilot_clsids) / sizeof(g_builtin_copilot_clsids[0]))},
    {L"set-desktop-background", L"Set as desktop background", BUILTIN_ACTION_REMOVE_PATHS, g_builtin_setdesktopwallpaper_paths, (int)(sizeof(g_builtin_setdesktopwallpaper_paths) / sizeof(g_builtin_setdesktopwallpaper_paths[0]))},
    {L"defender", L"Scan with Microsoft Defender", BUILTIN_ACTION_BLOCK_CLSID, g_builtin_defender_clsids, (int)(sizeof(g_builtin_defender_clsids) / sizeof(g_builtin_defender_clsids[0]))},
};

static int count_scopes(void) {
    return (int)(sizeof(g_scopes) / sizeof(g_scopes[0]));
}

static int count_builtin_targets(void) {
    return (int)(sizeof(g_builtin_targets) / sizeof(g_builtin_targets[0]));
}

static const wchar_t *g_default_options_filename = L"options.txt";

static wchar_t *load_text_file(const wchar_t *filename);
static int parse_import_record(const wchar_t *line, wchar_t *scope_name, DWORD scope_chars, wchar_t *entry_name, DWORD entry_chars, wchar_t *full_path, DWORD path_chars);
static int apply_builtin_target(const BUILTIN_TARGET *target);
static int preserved_export_contains_path(const PRESERVED_EXPORT_STATE *state, const wchar_t *path);

static int append_wchar(wchar_t *buffer, DWORD buffer_chars, DWORD *used, wchar_t ch) {
    if (*used + 1 >= buffer_chars) {
        return 0;
    }

    buffer[*used] = ch;
    *used += 1;
    buffer[*used] = L'\0';
    return 1;
}

static int append_text(wchar_t *buffer, DWORD buffer_chars, DWORD *used, const wchar_t *text) {
    const wchar_t *cursor;

    if (text == NULL) {
        return 1;
    }

    for (cursor = text; *cursor != L'\0'; cursor++) {
        if (!append_wchar(buffer, buffer_chars, used, *cursor)) {
            return 0;
        }
    }

    return 1;
}

static void print_last_error(const wchar_t *action, LONG code) {
    wchar_t message[512];
    DWORD flags;
    DWORD length;

    message[0] = L'\0';
    flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    length = FormatMessageW(flags, NULL, (DWORD)code, 0, message, (DWORD)(sizeof(message) / sizeof(message[0])), NULL);
    if (length == 0) {
        StringCchPrintfW(message, sizeof(message) / sizeof(message[0]), L"Windows error %ld", code);
    } else {
        while (length > 0 && (message[length - 1] == L'\r' || message[length - 1] == L'\n' || message[length - 1] == L' ' || message[length - 1] == L'\t')) {
            message[length - 1] = L'\0';
            length--;
        }
    }

    fwprintf(stderr, L"%ls failed: %ls\n", action, message);
}

static int query_registry_string(HKEY key, const wchar_t *value_name, wchar_t *buffer, DWORD buffer_chars) {
    BYTE raw[4096];
    DWORD raw_size;
    DWORD type;
    DWORD expanded;

    if (buffer == NULL || buffer_chars == 0) {
        return 0;
    }

    buffer[0] = L'\0';
    raw_size = (DWORD)sizeof(raw);
    type = 0;

    if (RegQueryValueExW(key, value_name, NULL, &type, raw, &raw_size) != ERROR_SUCCESS) {
        return 0;
    }

    if (type != REG_SZ && type != REG_EXPAND_SZ) {
        return 0;
    }

    ((wchar_t *)raw)[(sizeof(raw) / sizeof(wchar_t)) - 1] = L'\0';

    if (type == REG_EXPAND_SZ) {
        expanded = ExpandEnvironmentStringsW((const wchar_t *)raw, buffer, buffer_chars);
        if (expanded == 0 || expanded > buffer_chars) {
            return 0;
        }
        return 1;
    }

    if (FAILED(StringCchCopyW(buffer, buffer_chars, (const wchar_t *)raw))) {
        return 0;
    }

    return 1;
}

static int write_utf8_text(FILE *file, const wchar_t *text) {
    int utf8_bytes;
    char *utf8_buffer;
    size_t written;

    utf8_bytes = WideCharToMultiByte(CP_UTF8, 0, text, -1, NULL, 0, NULL, NULL);
    if (utf8_bytes <= 0) {
        return 0;
    }

    utf8_buffer = (char *)malloc((size_t)utf8_bytes);
    if (utf8_buffer == NULL) {
        return 0;
    }

    if (WideCharToMultiByte(CP_UTF8, 0, text, -1, utf8_buffer, utf8_bytes, NULL, NULL) <= 0) {
        free(utf8_buffer);
        return 0;
    }

    written = fwrite(utf8_buffer, 1, (size_t)(utf8_bytes - 1), file);
    free(utf8_buffer);
    return written == (size_t)(utf8_bytes - 1);
}

static int write_utf8_bom(FILE *file) {
    static const unsigned char bom[] = {0xEF, 0xBB, 0xBF};
    return fwrite(bom, 1, sizeof(bom), file) == sizeof(bom);
}

static wchar_t *duplicate_text(const wchar_t *text) {
    size_t chars;
    wchar_t *copy;

    if (text == NULL) {
        return NULL;
    }

    chars = wcslen(text);
    copy = (wchar_t *)malloc((chars + 1) * sizeof(wchar_t));
    if (copy == NULL) {
        return NULL;
    }

    if (FAILED(StringCchCopyW(copy, chars + 1, text))) {
        free(copy);
        return NULL;
    }

    return copy;
}

static int escape_field(const wchar_t *input, wchar_t *output, DWORD output_chars) {
    DWORD used;
    const wchar_t *cursor;

    if (output == NULL || output_chars == 0) {
        return 0;
    }

    output[0] = L'\0';
    used = 0;
    if (input == NULL) {
        return 1;
    }

    for (cursor = input; *cursor != L'\0'; cursor++) {
        wchar_t ch;

        ch = *cursor;
        if (ch == L'\\') {
            if (!append_text(output, output_chars, &used, L"\\\\")) {
                return 0;
            }
        } else if (ch == L'\t') {
            if (!append_text(output, output_chars, &used, L"\\t")) {
                return 0;
            }
        } else if (ch == L'\r') {
            if (!append_text(output, output_chars, &used, L"\\r")) {
                return 0;
            }
        } else if (ch == L'\n') {
            if (!append_text(output, output_chars, &used, L"\\n")) {
                return 0;
            }
        } else {
            if (!append_wchar(output, output_chars, &used, ch)) {
                return 0;
            }
        }
    }

    return 1;
}

static int unescape_field(const wchar_t *input, wchar_t *output, DWORD output_chars) {
    DWORD used;
    const wchar_t *cursor;

    if (output == NULL || output_chars == 0) {
        return 0;
    }

    output[0] = L'\0';
    used = 0;
    if (input == NULL) {
        return 1;
    }

    cursor = input;
    while (*cursor != L'\0') {
        if (*cursor == L'\\' && cursor[1] != L'\0') {
            wchar_t decoded;

            decoded = cursor[1];
            if (cursor[1] == L't') {
                decoded = L'\t';
            } else if (cursor[1] == L'r') {
                decoded = L'\r';
            } else if (cursor[1] == L'n') {
                decoded = L'\n';
            } else if (cursor[1] == L'\\') {
                decoded = L'\\';
            }

            if (!append_wchar(output, output_chars, &used, decoded)) {
                return 0;
            }

            cursor += 2;
            continue;
        }

        if (!append_wchar(output, output_chars, &used, *cursor)) {
            return 0;
        }
        cursor++;
    }

    return 1;
}

static void build_registry_path(const REGISTRY_LOCATION *location, const wchar_t *entry_name, wchar_t *buffer, DWORD buffer_chars) {
    if (entry_name != NULL && entry_name[0] != L'\0') {
        StringCchPrintfW(buffer, buffer_chars, L"%ls\\%ls\\%ls", location->root_name, location->subkey, entry_name);
    } else {
        StringCchPrintfW(buffer, buffer_chars, L"%ls\\%ls", location->root_name, location->subkey);
    }
}

static void load_entry_metadata(HKEY entry_key, wchar_t *display, DWORD display_chars, wchar_t *command, DWORD command_chars) {
    HKEY command_key;

    display[0] = L'\0';
    command[0] = L'\0';

    if (!query_registry_string(entry_key, L"MUIVerb", display, display_chars)) {
        if (!query_registry_string(entry_key, NULL, display, display_chars)) {
            if (!query_registry_string(entry_key, L"VerbName", display, display_chars)) {
                StringCchCopyW(display, display_chars, L"(no display text)");
            }
        }
    }

    if (RegOpenKeyExW(entry_key, L"command", 0, KEY_READ, &command_key) == ERROR_SUCCESS) {
        if (!query_registry_string(command_key, NULL, command, command_chars)) {
            command[0] = L'\0';
        }
        RegCloseKey(command_key);
    }

    if (command[0] == L'\0') {
        query_registry_string(entry_key, NULL, command, command_chars);
    }
}

static int contains_case_insensitive(const wchar_t *text, const wchar_t *filter) {
    const wchar_t *a;

    if (filter == NULL || filter[0] == L'\0') {
        return 1;
    }

    if (text == NULL) {
        return 0;
    }

    for (a = text; *a != L'\0'; a++) {
        const wchar_t *x = a;
        const wchar_t *y = filter;

        while (*x != L'\0' && *y != L'\0' && towlower(*x) == towlower(*y)) {
            x++;
            y++;
        }

        if (*y == L'\0') {
            return 1;
        }
    }

    return 0;
}

static int copy_field_token(const wchar_t *start, wchar_t delimiter, wchar_t *buffer, DWORD buffer_chars, const wchar_t **next) {
    DWORD used;
    const wchar_t *cursor;

    if (buffer == NULL || buffer_chars == 0) {
        return 0;
    }

    buffer[0] = L'\0';
    used = 0;
    cursor = start;
    while (*cursor != L'\0' && *cursor != delimiter) {
        if (!append_wchar(buffer, buffer_chars, &used, *cursor)) {
            return 0;
        }
        cursor++;
    }

    if (next != NULL) {
        *next = cursor;
        if (**next == delimiter) {
            *next += 1;
        }
    }

    return 1;
}

static int process_is_elevated(void) {
    SID_IDENTIFIER_AUTHORITY authority;
    PSID administrators_group;
    BOOL is_member;

    authority = SECURITY_NT_AUTHORITY;
    administrators_group = NULL;
    is_member = FALSE;

    if (!AllocateAndInitializeSid(&authority, 2,
            SECURITY_BUILTIN_DOMAIN_RID,
            DOMAIN_ALIAS_RID_ADMINS,
            0, 0, 0, 0, 0, 0,
            &administrators_group)) {
        return 0;
    }

    if (!CheckTokenMembership(NULL, administrators_group, &is_member)) {
        is_member = FALSE;
    }

    FreeSid(administrators_group);
    return is_member ? 1 : 0;
}

static int entry_matches_filter(const wchar_t *filter, const wchar_t *entry_name, const wchar_t *display, const wchar_t *command, const wchar_t *full_path) {
    if (filter == NULL || filter[0] == L'\0') {
        return 1;
    }

    if (contains_case_insensitive(entry_name, filter)) {
        return 1;
    }

    if (contains_case_insensitive(display, filter)) {
        return 1;
    }

    if (contains_case_insensitive(command, filter)) {
        return 1;
    }

    if (contains_case_insensitive(full_path, filter)) {
        return 1;
    }

    return 0;
}

static const SCOPE *find_scope(const wchar_t *name) {
    int i;

    if (name == NULL) {
        return NULL;
    }

    for (i = 0; i < count_scopes(); i++) {
        if (_wcsicmp(g_scopes[i].name, name) == 0) {
            return &g_scopes[i];
        }
    }

    return NULL;
}

static const BUILTIN_TARGET *find_builtin_target(const wchar_t *query) {
    int i;

    if (query == NULL) {
        return NULL;
    }

    for (i = 0; i < count_builtin_targets(); i++) {
        if (_wcsicmp(g_builtin_targets[i].name, query) == 0 || _wcsicmp(g_builtin_targets[i].label, query) == 0) {
            return &g_builtin_targets[i];
        }
    }

    return NULL;
}

static int is_builtin_scope_name(const wchar_t *scope_name) {
    return scope_name != NULL && _wcsicmp(scope_name, g_builtin_scope_name) == 0;
}

static int build_builtin_identifier(const BUILTIN_TARGET *target, wchar_t *buffer, DWORD buffer_chars) {
    if (target == NULL || buffer == NULL || buffer_chars == 0) {
        return 0;
    }

    return SUCCEEDED(StringCchPrintfW(buffer, buffer_chars, L"%ls%ls", g_builtin_path_prefix, target->name));
}

static const BUILTIN_TARGET *find_builtin_target_from_identifier(const wchar_t *identifier) {
    size_t prefix_chars;

    if (identifier == NULL) {
        return NULL;
    }

    prefix_chars = wcslen(g_builtin_path_prefix);
    if (_wcsnicmp(identifier, g_builtin_path_prefix, prefix_chars) != 0) {
        return NULL;
    }

    return find_builtin_target(identifier + prefix_chars);
}

static int split_registry_path(const wchar_t *full_path, HKEY *root_key, const wchar_t **subkey) {
    if (full_path == NULL || root_key == NULL || subkey == NULL) {
        return 0;
    }

    if (_wcsnicmp(full_path, L"HKCU\\", 5) == 0) {
        *root_key = HKEY_CURRENT_USER;
        *subkey = full_path + 5;
        return 1;
    }

    if (_wcsnicmp(full_path, L"HKLM\\", 5) == 0) {
        *root_key = HKEY_LOCAL_MACHINE;
        *subkey = full_path + 5;
        return 1;
    }

    return 0;
}

static void print_usage(void) {
    int i;

    wprintf(L"rcmenuedit: inspect and remove Windows context-menu entries\n");
    wprintf(L"\n");
    wprintf(L"Usage:\n");
    wprintf(L"  rcmenuedit help\n");
    wprintf(L"  rcmenuedit scopes\n");
    wprintf(L"  rcmenuedit known [filter]\n");
    wprintf(L"  rcmenuedit all [filter]\n");
    wprintf(L"  rcmenuedit builtins [filter]\n");
    wprintf(L"  rcmenuedit exportoptions [file]\n");
    wprintf(L"  rcmenuedit importoptions [file]\n");
    wprintf(L"  rcmenuedit find [scope] [filter]\n");
    wprintf(L"  rcmenuedit list [scope] [filter]\n");
    wprintf(L"  rcmenuedit remove <scope> <entry-name-or-label>\n");
    wprintf(L"  rcmenuedit removebuiltin <name-or-label>\n");
    wprintf(L"\n");
    wprintf(L"Scopes:\n");
    for (i = 0; i < count_scopes(); i++) {
        wprintf(L"  %-8ls %ls\n", g_scopes[i].name, g_scopes[i].description);
    }
    wprintf(L"\n");
    wprintf(L"Built-in targets:\n");
    for (i = 0; i < count_builtin_targets(); i++) {
        wprintf(L"  %-24ls %ls\n", g_builtin_targets[i].name, g_builtin_targets[i].label);
    }
    wprintf(L"\n");
    wprintf(L"Examples:\n");
    wprintf(L"  rcmenuedit known\n");
    wprintf(L"  rcmenuedit builtins\n");
    wprintf(L"  rcmenuedit all terminal\n");
    wprintf(L"  rcmenuedit exportoptions options.txt\n");
    wprintf(L"  rcmenuedit exportoptions\n");
    wprintf(L"  rcmenuedit importoptions options.txt\n");
    wprintf(L"  rcmenuedit importoptions\n");
    wprintf(L"  rcmenuedit find\n");
    wprintf(L"  rcmenuedit find file 7-zip\n");
    wprintf(L"  rcmenuedit remove desktop \"Windows Terminal\"\n");
    wprintf(L"  rcmenuedit removebuiltin \"Ask Copilot\"\n");
    wprintf(L"\n");
    wprintf(L"When [file] is omitted for exportoptions/importoptions, %ls is used.\n", g_default_options_filename);
}

static void print_scopes(void) {
    int i;

    for (i = 0; i < count_scopes(); i++) {
        wprintf(L"%ls: %ls\n", g_scopes[i].name, g_scopes[i].description);
    }
}

static int enumerate_location_entries(const REGISTRY_LOCATION *location, const wchar_t *filter, ENTRY_CALLBACK callback, void *context) {
    HKEY location_key;
    LONG result;
    DWORD index;
    int found;

    result = RegOpenKeyExW(location->root_key, location->subkey, 0, KEY_READ, &location_key);
    if (result == ERROR_FILE_NOT_FOUND) {
        return 0;
    }

    if (result != ERROR_SUCCESS) {
        wchar_t path[512];
        build_registry_path(location, NULL, path, sizeof(path) / sizeof(path[0]));
        print_last_error(path, result);
        return 0;
    }

    found = 0;
    index = 0;
    for (;;) {
        wchar_t entry_name[260];
        DWORD entry_name_chars;
        HKEY entry_key;
        LONG enum_result;

        entry_name_chars = (DWORD)(sizeof(entry_name) / sizeof(entry_name[0]));
        enum_result = RegEnumKeyExW(location_key, index, entry_name, &entry_name_chars, NULL, NULL, NULL, NULL);
        if (enum_result == ERROR_NO_MORE_ITEMS) {
            break;
        }

        if (enum_result != ERROR_SUCCESS) {
            print_last_error(L"RegEnumKeyExW", enum_result);
            break;
        }

        result = RegOpenKeyExW(location_key, entry_name, 0, KEY_READ, &entry_key);
        if (result == ERROR_SUCCESS) {
            wchar_t display[1024];
            wchar_t command[2048];
            wchar_t path[1024];

            load_entry_metadata(entry_key, display, sizeof(display) / sizeof(display[0]), command, sizeof(command) / sizeof(command[0]));
            build_registry_path(location, entry_name, path, sizeof(path) / sizeof(path[0]));

            if (entry_matches_filter(filter, entry_name, display, command, path)) {
                if (callback != NULL) {
                    callback(location, entry_name, display, command, path, context);
                }
                found++;
            }

            RegCloseKey(entry_key);
        }

        index++;
    }

    RegCloseKey(location_key);
    return found;
}

static int print_entry_callback(const REGISTRY_LOCATION *location, const wchar_t *entry_name, const wchar_t *display, const wchar_t *command, const wchar_t *path, void *context) {
    (void)context;

    wprintf(L"scope=%ls type=%ls key=%ls label=%ls path=%ls",
        location->scope_name,
        location->entry_type,
        entry_name,
        display,
        path);
    if (command[0] != L'\0') {
        wprintf(L" value=%ls", command);
    }
    wprintf(L"\n");

    return 1;
}

static int list_location_entries(const REGISTRY_LOCATION *location, const wchar_t *filter) {
    return enumerate_location_entries(location, filter, print_entry_callback, NULL);
}

static int list_scope_entries(const SCOPE *scope, const wchar_t *filter) {
    int i;
    int total;

    total = 0;
    for (i = 0; i < scope->location_count; i++) {
        total += list_location_entries(&scope->locations[i], filter);
    }
    return total;
}

static int run_find(const SCOPE *scope, const wchar_t *filter) {
    int i;
    int total;

    total = 0;
    if (scope != NULL) {
        total = list_scope_entries(scope, filter);
    } else {
        for (i = 0; i < count_scopes(); i++) {
            total += list_scope_entries(&g_scopes[i], filter);
        }
    }

    if (total == 0) {
        wprintf(L"No matching entries found.\n");
    }

    return 0;
}

static int run_builtins(const wchar_t *filter) {
    int i;
    int total;

    total = 0;
    for (i = 0; i < count_builtin_targets(); i++) {
        const BUILTIN_TARGET *target;
        const wchar_t *action_name;

        target = &g_builtin_targets[i];
        if (!contains_case_insensitive(target->name, filter) && !contains_case_insensitive(target->label, filter)) {
            continue;
        }

        action_name = target->action_kind == BUILTIN_ACTION_BLOCK_CLSID ? L"block-clsid" : L"remove-path";
        wprintf(L"name=%ls label=%ls action=%ls targets=%d\n",
            target->name,
            target->label,
            action_name,
            target->value_count);
        total++;
    }

    if (total == 0) {
        wprintf(L"No matching built-in targets found.\n");
    }

    return 0;
}

static int export_builtin_entry(const BUILTIN_TARGET *target, EXPORT_CONTEXT *export_context) {
    wchar_t builtin_identifier[256];
    wchar_t escaped_scope[128];
    wchar_t escaped_name[256];
    wchar_t escaped_path[256];
    wchar_t escaped_label[512];
    wchar_t line[2048];

    if (target == NULL || export_context == NULL || export_context->file == NULL) {
        return 0;
    }

    if (!build_builtin_identifier(target, builtin_identifier, sizeof(builtin_identifier) / sizeof(builtin_identifier[0]))) {
        export_context->write_failed = 1;
        return 0;
    }

    if (preserved_export_contains_path(export_context->preserved_state, builtin_identifier)) {
        export_context->skipped_preserved_entries++;
        return 1;
    }

    if (!escape_field(g_builtin_scope_name, escaped_scope, sizeof(escaped_scope) / sizeof(escaped_scope[0])) ||
        !escape_field(target->name, escaped_name, sizeof(escaped_name) / sizeof(escaped_name[0])) ||
        !escape_field(builtin_identifier, escaped_path, sizeof(escaped_path) / sizeof(escaped_path[0])) ||
        !escape_field(target->label, escaped_label, sizeof(escaped_label) / sizeof(escaped_label[0]))) {
        export_context->write_failed = 1;
        return 0;
    }

    if (FAILED(StringCchPrintfW(line, sizeof(line) / sizeof(line[0]), L"k %ls\t%ls\t%ls\t%ls\r\n", escaped_scope, escaped_name, escaped_path, escaped_label))) {
        export_context->write_failed = 1;
        return 0;
    }

    if (!write_utf8_text(export_context->file, line)) {
        export_context->write_failed = 1;
        return 0;
    }

    export_context->written_k_entries++;
    return 1;
}

static void free_preserved_export_state(PRESERVED_EXPORT_STATE *state) {
    int i;

    if (state == NULL) {
        return;
    }

    for (i = 0; i < state->count; i++) {
        free(state->entries[i].line);
        free(state->entries[i].path);
    }

    free(state->entries);
    state->entries = NULL;
    state->count = 0;
    state->capacity = 0;
    state->malformed_count = 0;
}

static int append_preserved_export_entry(PRESERVED_EXPORT_STATE *state, const wchar_t *line, const wchar_t *path) {
    PRESERVED_EXPORT_ENTRY *expanded_entries;
    int new_capacity;

    if (state == NULL || line == NULL) {
        return 0;
    }

    if (state->count == state->capacity) {
        new_capacity = state->capacity == 0 ? 16 : state->capacity * 2;
        expanded_entries = (PRESERVED_EXPORT_ENTRY *)realloc(state->entries, (size_t)new_capacity * sizeof(PRESERVED_EXPORT_ENTRY));
        if (expanded_entries == NULL) {
            return 0;
        }

        state->entries = expanded_entries;
        state->capacity = new_capacity;
    }

    state->entries[state->count].line = duplicate_text(line);
    if (state->entries[state->count].line == NULL) {
        return 0;
    }

    state->entries[state->count].path = duplicate_text(path);
    if (path != NULL && state->entries[state->count].path == NULL) {
        free(state->entries[state->count].line);
        state->entries[state->count].line = NULL;
        return 0;
    }

    state->count++;
    return 1;
}

static int preserved_export_contains_path(const PRESERVED_EXPORT_STATE *state, const wchar_t *path) {
    int i;

    if (state == NULL || path == NULL || path[0] == L'\0') {
        return 0;
    }

    for (i = 0; i < state->count; i++) {
        if (state->entries[i].path != NULL && _wcsicmp(state->entries[i].path, path) == 0) {
            return 1;
        }
    }

    return 0;
}

static int write_preserved_export_entries(FILE *file, const PRESERVED_EXPORT_STATE *state) {
    int i;

    if (file == NULL || state == NULL) {
        return 1;
    }

    for (i = 0; i < state->count; i++) {
        if (!write_utf8_text(file, state->entries[i].line) || !write_utf8_text(file, L"\r\n")) {
            return 0;
        }
    }

    return 1;
}

static int load_preserved_export_state(const wchar_t *filename, PRESERVED_EXPORT_STATE *state) {
    DWORD attributes;
    wchar_t *text;
    wchar_t *cursor;

    if (state == NULL) {
        return 0;
    }

    state->entries = NULL;
    state->count = 0;
    state->capacity = 0;
    state->malformed_count = 0;

    attributes = GetFileAttributesW(filename);
    if (attributes == INVALID_FILE_ATTRIBUTES) {
        DWORD error_code;

        error_code = GetLastError();
        if (error_code == ERROR_FILE_NOT_FOUND || error_code == ERROR_PATH_NOT_FOUND) {
            return 1;
        }

        print_last_error(filename, (LONG)error_code);
        return 0;
    }

    text = load_text_file(filename);
    if (text == NULL) {
        return 0;
    }

    cursor = text;
    while (*cursor != L'\0') {
        wchar_t *line_start;
        wchar_t *line_end;
        wchar_t saved_char;

        line_start = cursor;
        line_end = cursor;
        while (*line_end != L'\0' && *line_end != L'\r' && *line_end != L'\n') {
            line_end++;
        }

        saved_char = *line_end;
        *line_end = L'\0';

        if (wcsncmp(line_start, L"r ", 2) == 0) {
            wchar_t scope_name[256];
            wchar_t entry_name[512];
            wchar_t full_path[2048];
            int parse_result;
            const wchar_t *path_to_store;

            path_to_store = NULL;
            parse_result = parse_import_record(line_start, scope_name, sizeof(scope_name) / sizeof(scope_name[0]), entry_name, sizeof(entry_name) / sizeof(entry_name[0]), full_path, sizeof(full_path) / sizeof(full_path[0]));
            if (parse_result > 0) {
                path_to_store = full_path;
            } else {
                state->malformed_count++;
            }

            if (!append_preserved_export_entry(state, line_start, path_to_store)) {
                free(text);
                free_preserved_export_state(state);
                fwprintf(stderr, L"Out of memory preserving existing remove entries from %ls.\n", filename);
                return 0;
            }
        }

        *line_end = saved_char;
        if (*line_end == L'\r' && line_end[1] == L'\n') {
            cursor = line_end + 2;
        } else if (*line_end == L'\r' || *line_end == L'\n') {
            cursor = line_end + 1;
        } else {
            cursor = line_end;
        }
    }

    free(text);
    return 1;
}

static int export_entry_callback(const REGISTRY_LOCATION *location, const wchar_t *entry_name, const wchar_t *display, const wchar_t *command, const wchar_t *path, void *context) {
    EXPORT_CONTEXT *export_context;
    wchar_t escaped_scope[128];
    wchar_t escaped_key[512];
    wchar_t escaped_path[2048];
    wchar_t escaped_label[2048];
    wchar_t line[8192];

    (void)command;

    export_context = (EXPORT_CONTEXT *)context;
    if (export_context == NULL || export_context->file == NULL) {
        return 0;
    }

    if (preserved_export_contains_path(export_context->preserved_state, path)) {
        export_context->skipped_preserved_entries++;
        return 1;
    }

    if (!escape_field(location->scope_name, escaped_scope, sizeof(escaped_scope) / sizeof(escaped_scope[0])) ||
        !escape_field(entry_name, escaped_key, sizeof(escaped_key) / sizeof(escaped_key[0])) ||
        !escape_field(path, escaped_path, sizeof(escaped_path) / sizeof(escaped_path[0])) ||
        !escape_field(display, escaped_label, sizeof(escaped_label) / sizeof(escaped_label[0]))) {
        export_context->write_failed = 1;
        return 0;
    }

    if (FAILED(StringCchPrintfW(line, sizeof(line) / sizeof(line[0]), L"k %ls\t%ls\t%ls\t%ls\r\n", escaped_scope, escaped_key, escaped_path, escaped_label))) {
        export_context->write_failed = 1;
        return 0;
    }

    if (!write_utf8_text(export_context->file, line)) {
        export_context->write_failed = 1;
        return 0;
    }

    export_context->written_k_entries++;
    return 1;
}

static int run_exportoptions(const wchar_t *filename) {
    FILE *file;
    EXPORT_CONTEXT export_context;
    PRESERVED_EXPORT_STATE preserved_state;
    int i;

    if (!load_preserved_export_state(filename, &preserved_state)) {
        return 1;
    }

    file = NULL;
    if (_wfopen_s(&file, filename, L"wb") != 0 || file == NULL) {
        fwprintf(stderr, L"Unable to open %ls for writing.\n", filename);
        free_preserved_export_state(&preserved_state);
        return 1;
    }

    if (!write_utf8_bom(file)) {
        fwprintf(stderr, L"Unable to write BOM to %ls.\n", filename);
        fclose(file);
        free_preserved_export_state(&preserved_state);
        return 1;
    }

    if (!write_preserved_export_entries(file, &preserved_state)) {
        fwprintf(stderr, L"Unable to preserve existing remove entries in %ls.\n", filename);
        fclose(file);
        free_preserved_export_state(&preserved_state);
        return 1;
    }

    export_context.file = file;
    export_context.write_failed = 0;
    export_context.preserved_state = &preserved_state;
    export_context.written_k_entries = 0;
    export_context.skipped_preserved_entries = 0;
    for (i = 0; i < count_scopes(); i++) {
        {
            int j;
            for (j = 0; j < g_scopes[i].location_count; j++) {
                enumerate_location_entries(&g_scopes[i].locations[j], NULL, export_entry_callback, &export_context);
            }
        }
    }
    for (i = 0; i < count_builtin_targets(); i++) {
        if (!export_builtin_entry(&g_builtin_targets[i], &export_context)) {
            break;
        }
    }

    fclose(file);

    if (export_context.write_failed) {
        fwprintf(stderr, L"Export failed while writing %ls.\n", filename);
        free_preserved_export_state(&preserved_state);
        return 1;
    }

    wprintf(L"Exported %d keep entries to %ls\n", export_context.written_k_entries, filename);
    if (preserved_state.count > 0) {
        wprintf(L"Preserved %d existing remove entr%s.\n", preserved_state.count, preserved_state.count == 1 ? L"y" : L"ies");
    }
    if (export_context.skipped_preserved_entries > 0) {
        wprintf(L"Skipped %d keep entr%s already marked for removal.\n", export_context.skipped_preserved_entries, export_context.skipped_preserved_entries == 1 ? L"y" : L"ies");
    }
    if (preserved_state.malformed_count > 0) {
        fwprintf(stderr, L"warning: preserved %d malformed remove line%s as-is.\n", preserved_state.malformed_count, preserved_state.malformed_count == 1 ? L"" : L"s");
    }
    wprintf(L"Change leading \"k \" markers to \"r \" for entries you want importoptions to remove.\n");
    free_preserved_export_state(&preserved_state);
    return 0;
}

static wchar_t *load_text_file(const wchar_t *filename) {
    FILE *file;
    long size;
    unsigned char *bytes;
    wchar_t *text;
    const unsigned char *content;
    int content_bytes;
    int wide_chars;
    UINT code_page;
    DWORD flags;

    file = NULL;
    if (_wfopen_s(&file, filename, L"rb") != 0 || file == NULL) {
        fwprintf(stderr, L"Unable to open %ls for reading.\n", filename);
        return NULL;
    }

    if (fseek(file, 0, SEEK_END) != 0) {
        fclose(file);
        fwprintf(stderr, L"Unable to seek in %ls.\n", filename);
        return NULL;
    }

    size = ftell(file);
    if (size < 0) {
        fclose(file);
        fwprintf(stderr, L"Unable to get file size for %ls.\n", filename);
        return NULL;
    }

    if (fseek(file, 0, SEEK_SET) != 0) {
        fclose(file);
        fwprintf(stderr, L"Unable to rewind %ls.\n", filename);
        return NULL;
    }

    bytes = (unsigned char *)malloc((size_t)size + 1);
    if (bytes == NULL) {
        fclose(file);
        fwprintf(stderr, L"Out of memory reading %ls.\n", filename);
        return NULL;
    }

    if (size > 0 && fread(bytes, 1, (size_t)size, file) != (size_t)size) {
        fclose(file);
        free(bytes);
        fwprintf(stderr, L"Unable to read %ls.\n", filename);
        return NULL;
    }

    fclose(file);
    bytes[size] = 0;

    content = bytes;
    content_bytes = (int)size;
    if (content_bytes >= 3 && content[0] == 0xEF && content[1] == 0xBB && content[2] == 0xBF) {
        content += 3;
        content_bytes -= 3;
    }

    code_page = CP_UTF8;
    flags = MB_ERR_INVALID_CHARS;
    wide_chars = MultiByteToWideChar(code_page, flags, (const char *)content, content_bytes, NULL, 0);
    if (wide_chars <= 0) {
        code_page = CP_ACP;
        flags = 0;
        wide_chars = MultiByteToWideChar(code_page, flags, (const char *)content, content_bytes, NULL, 0);
        if (wide_chars <= 0) {
            free(bytes);
            fwprintf(stderr, L"Unable to decode %ls.\n", filename);
            return NULL;
        }
    }

    text = (wchar_t *)malloc(((size_t)wide_chars + 1) * sizeof(wchar_t));
    if (text == NULL) {
        free(bytes);
        fwprintf(stderr, L"Out of memory decoding %ls.\n", filename);
        return NULL;
    }

    if (MultiByteToWideChar(code_page, flags, (const char *)content, content_bytes, text, wide_chars) <= 0) {
        free(bytes);
        free(text);
        fwprintf(stderr, L"Unable to convert %ls.\n", filename);
        return NULL;
    }

    text[wide_chars] = L'\0';
    free(bytes);
    return text;
}

static int remove_registry_path_if_exists(const wchar_t *full_path) {
    HKEY root_key;
    const wchar_t *subkey;
    HKEY key;
    LONG result;

    if (!split_registry_path(full_path, &root_key, &subkey)) {
        fwprintf(stderr, L"Unsupported registry path: %ls\n", full_path);
        return -1;
    }

    result = RegOpenKeyExW(root_key, subkey, 0, KEY_READ, &key);
    if (result == ERROR_FILE_NOT_FOUND) {
        return 0;
    }

    if (result != ERROR_SUCCESS) {
        print_last_error(full_path, result);
        return -1;
    }

    RegCloseKey(key);

    result = RegDeleteTreeW(root_key, subkey);
    if (result != ERROR_SUCCESS) {
        print_last_error(full_path, result);
        return -1;
    }

    wprintf(L"Removed %ls\n", full_path);
    return 1;
}

static int parse_import_record(const wchar_t *line, wchar_t *scope_name, DWORD scope_chars, wchar_t *entry_name, DWORD entry_chars, wchar_t *full_path, DWORD path_chars) {
    const wchar_t *cursor;
    wchar_t raw_scope[256];
    wchar_t raw_key[512];
    wchar_t raw_path[2048];

    if (line == NULL || wcsncmp(line, L"r ", 2) != 0) {
        return 0;
    }

    cursor = line + 2;
    if (!copy_field_token(cursor, L'\t', raw_scope, sizeof(raw_scope) / sizeof(raw_scope[0]), &cursor)) {
        return -1;
    }
    if (!copy_field_token(cursor, L'\t', raw_key, sizeof(raw_key) / sizeof(raw_key[0]), &cursor)) {
        return -1;
    }
    if (!copy_field_token(cursor, L'\t', raw_path, sizeof(raw_path) / sizeof(raw_path[0]), &cursor)) {
        return -1;
    }

    if (!unescape_field(raw_scope, scope_name, scope_chars) ||
        !unescape_field(raw_key, entry_name, entry_chars) ||
        !unescape_field(raw_path, full_path, path_chars)) {
        return -1;
    }

    if (scope_name[0] == L'\0' || entry_name[0] == L'\0' || full_path[0] == L'\0') {
        return -1;
    }

    return 1;
}

static int run_importoptions(const wchar_t *filename) {
    wchar_t *text;
    wchar_t *cursor;
    int line_number;
    int requested;
    int removed;
    int errors;

    text = load_text_file(filename);
    if (text == NULL) {
        return 1;
    }

    if (!process_is_elevated()) {
        fwprintf(stderr, L"warning: process is not elevated; machine-wide keys may fail with access denied.\n");
    }

    cursor = text;
    line_number = 1;
    requested = 0;
    removed = 0;
    errors = 0;

    while (*cursor != L'\0') {
        wchar_t *line_start;
        wchar_t *line_end;
        wchar_t saved_char;

        line_start = cursor;
        line_end = cursor;
        while (*line_end != L'\0' && *line_end != L'\r' && *line_end != L'\n') {
            line_end++;
        }

        saved_char = *line_end;
        *line_end = L'\0';

        if (wcsncmp(line_start, L"r ", 2) == 0) {
            wchar_t scope_name[256];
            wchar_t entry_name[512];
            wchar_t full_path[2048];
            int parse_result;
            int remove_result;

            parse_result = parse_import_record(line_start, scope_name, sizeof(scope_name) / sizeof(scope_name[0]), entry_name, sizeof(entry_name) / sizeof(entry_name[0]), full_path, sizeof(full_path) / sizeof(full_path[0]));
            requested++;
            if (parse_result <= 0) {
                fwprintf(stderr, L"Malformed import line %d\n", line_number);
                errors++;
            } else {
                if (is_builtin_scope_name(scope_name)) {
                    const BUILTIN_TARGET *target;

                    target = find_builtin_target_from_identifier(full_path);
                    if (target == NULL) {
                        target = find_builtin_target(entry_name);
                    }

                    if (target == NULL) {
                        fwprintf(stderr, L"Unknown built-in target on line %d: %ls\n", line_number, entry_name);
                        errors++;
                    } else {
                        remove_result = apply_builtin_target(target);
                        if (remove_result >= 0) {
                            if (remove_result > 0) {
                                removed++;
                            }
                        } else {
                            fwprintf(stderr, L"Failed to remove line %d (%ls / %ls)\n", line_number, scope_name, entry_name);
                            errors++;
                        }
                    }
                } else {
                    const SCOPE *scope;

                    scope = find_scope(scope_name);
                    if (scope == NULL) {
                        fwprintf(stderr, L"Unknown scope on line %d: %ls\n", line_number, scope_name);
                        errors++;
                    } else {
                        remove_result = remove_registry_path_if_exists(full_path);
                        if (remove_result > 0) {
                            removed++;
                        } else if (remove_result < 0) {
                            fwprintf(stderr, L"Failed to remove line %d (%ls / %ls)\n", line_number, scope_name, entry_name);
                            errors++;
                        }
                    }
                }
            }
        }

        *line_end = saved_char;
        if (*line_end == L'\r' && line_end[1] == L'\n') {
            cursor = line_end + 2;
        } else if (*line_end == L'\r' || *line_end == L'\n') {
            cursor = line_end + 1;
        } else {
            cursor = line_end;
        }
        line_number++;
    }

    free(text);

    wprintf(L"Import processed %d marked entries and removed %d.\n", requested, removed);
    if (errors > 0) {
        fwprintf(stderr, L"Import finished with %d error%s.\n", errors, errors == 1 ? L"" : L"s");
        return 1;
    }

    return 0;
}

static int block_shell_extension_clsid(const wchar_t *clsid) {
    static const wchar_t *blocked_subkey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Blocked";
    HKEY blocked_key;
    LONG result;
    int already_blocked;

    blocked_key = NULL;
    result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, blocked_subkey, 0, NULL, 0, KEY_QUERY_VALUE | KEY_SET_VALUE, NULL, &blocked_key, NULL);
    if (result != ERROR_SUCCESS) {
        print_last_error(blocked_subkey, result);
        return -1;
    }

    already_blocked = RegQueryValueExW(blocked_key, clsid, NULL, NULL, NULL, NULL) == ERROR_SUCCESS;
    result = RegSetValueExW(blocked_key, clsid, 0, REG_SZ, (const BYTE *)L"", (DWORD)sizeof(wchar_t));
    RegCloseKey(blocked_key);

    if (result != ERROR_SUCCESS) {
        print_last_error(clsid, result);
        return -1;
    }

    if (already_blocked) {
        wprintf(L"Already blocked %ls\n", clsid);
        return 0;
    }

    wprintf(L"Blocked %ls\n", clsid);
    return 1;
}

static int apply_builtin_target(const BUILTIN_TARGET *target) {
    int i;
    int changes;
    int errors;

    if (target == NULL) {
        return -1;
    }

    changes = 0;
    errors = 0;
    for (i = 0; i < target->value_count; i++) {
        int result;

        if (target->action_kind == BUILTIN_ACTION_BLOCK_CLSID) {
            result = block_shell_extension_clsid(target->values[i]);
        } else {
            result = remove_registry_path_if_exists(target->values[i]);
        }

        if (result > 0) {
            changes++;
        } else if (result < 0) {
            errors++;
        }
    }

    if (errors > 0) {
        return -1;
    }

    return changes;
}

static int run_removebuiltin(const BUILTIN_TARGET *target) {
    int changes;

    if (target == NULL) {
        fwprintf(stderr, L"removebuiltin requires a valid built-in target.\n");
        return 1;
    }

    if (!process_is_elevated()) {
        fwprintf(stderr, L"warning: process is not elevated; built-in changes under HKLM may fail with access denied.\n");
    }

    changes = apply_builtin_target(target);
    if (changes < 0) {
        fwprintf(stderr, L"Built-in target %ls finished with errors.\n", target->label);
        return 1;
    }

    if (changes == 0) {
        wprintf(L"No changes were needed for %ls.\n", target->label);
    } else {
        wprintf(L"Applied %d change%s for %ls.\n", changes, changes == 1 ? L"" : L"s", target->label);
    }

    return 0;
}

static int entry_matches_remove_target(const wchar_t *query, const wchar_t *entry_name, const wchar_t *display) {
    if (_wcsicmp(query, entry_name) == 0) {
        return 1;
    }

    if (display != NULL && display[0] != L'\0' && _wcsicmp(query, display) == 0) {
        return 1;
    }

    return 0;
}

static int remove_matches_in_location(const REGISTRY_LOCATION *location, const wchar_t *query) {
    int removed;

    removed = 0;
    for (;;) {
        HKEY location_key;
        LONG open_result;
        DWORD index;
        int removed_this_pass;

        open_result = RegOpenKeyExW(location->root_key, location->subkey, 0, KEY_READ | KEY_WRITE, &location_key);
        if (open_result == ERROR_FILE_NOT_FOUND) {
            return removed;
        }

        if (open_result != ERROR_SUCCESS) {
            wchar_t path[512];
            build_registry_path(location, NULL, path, sizeof(path) / sizeof(path[0]));
            print_last_error(path, open_result);
            return removed;
        }

        removed_this_pass = 0;
        index = 0;
        for (;;) {
            wchar_t entry_name[260];
            DWORD entry_name_chars;
            HKEY entry_key;
            LONG enum_result;

            entry_name_chars = (DWORD)(sizeof(entry_name) / sizeof(entry_name[0]));
            enum_result = RegEnumKeyExW(location_key, index, entry_name, &entry_name_chars, NULL, NULL, NULL, NULL);
            if (enum_result == ERROR_NO_MORE_ITEMS) {
                break;
            }

            if (enum_result != ERROR_SUCCESS) {
                print_last_error(L"RegEnumKeyExW", enum_result);
                break;
            }

            if (RegOpenKeyExW(location_key, entry_name, 0, KEY_READ, &entry_key) == ERROR_SUCCESS) {
                wchar_t display[1024];
                wchar_t command[2048];

                load_entry_metadata(entry_key, display, sizeof(display) / sizeof(display[0]), command, sizeof(command) / sizeof(command[0]));
                RegCloseKey(entry_key);

                if (entry_matches_remove_target(query, entry_name, display)) {
                    wchar_t full_path[1024];
                    LONG delete_result;

                    build_registry_path(location, entry_name, full_path, sizeof(full_path) / sizeof(full_path[0]));
                    delete_result = RegDeleteTreeW(location_key, entry_name);
                    if (delete_result == ERROR_SUCCESS) {
                        wprintf(L"Removed %ls\n", full_path);
                        removed++;
                        removed_this_pass = 1;
                    } else {
                        print_last_error(full_path, delete_result);
                    }
                    break;
                }
            }

            index++;
        }

        RegCloseKey(location_key);
        if (!removed_this_pass) {
            break;
        }
    }

    return removed;
}

static int run_remove(const SCOPE *scope, const wchar_t *query) {
    int i;
    int removed;

    if (scope == NULL || query == NULL || query[0] == L'\0') {
        fwprintf(stderr, L"remove requires a valid scope and entry name.\n");
        return 1;
    }

    if (!process_is_elevated()) {
        fwprintf(stderr, L"warning: process is not elevated; machine-wide keys may fail with access denied.\n");
    }

    removed = 0;
    for (i = 0; i < scope->location_count; i++) {
        removed += remove_matches_in_location(&scope->locations[i], query);
    }

    if (removed == 0) {
        fwprintf(stderr, L"No entries matched %ls in scope %ls.\n", query, scope->name);
        return 1;
    }

    wprintf(L"Removed %d entr%s.\n", removed, removed == 1 ? L"y" : L"ies");
    return 0;
}

static int command_is(const wchar_t *actual, const wchar_t *expected) {
    return _wcsicmp(actual, expected) == 0;
}

int main(void) {
    LPWSTR *argv;
    int argc;
    const wchar_t *command;

    argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv == NULL) {
        print_last_error(L"CommandLineToArgvW", GetLastError());
        return 1;
    }

    if (argc < 2) {
        print_usage();
        LocalFree(argv);
        return 0;
    }

    command = argv[1];

    if (command_is(command, L"help") || command_is(command, L"/?") || command_is(command, L"-?") || command_is(command, L"--help")) {
        print_usage();
        LocalFree(argv);
        return 0;
    }

    if (command_is(command, L"scopes")) {
        print_scopes();
        LocalFree(argv);
        return 0;
    }

    if (command_is(command, L"known") || command_is(command, L"all")) {
        const wchar_t *filter;
        int exit_code;

        filter = NULL;
        if (argc >= 3) {
            filter = argv[2];
        }

        exit_code = run_find(NULL, filter);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"builtins")) {
        const wchar_t *filter;
        int exit_code;

        filter = NULL;
        if (argc >= 3) {
            filter = argv[2];
        }

        exit_code = run_builtins(filter);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"exportoptions")) {
        int exit_code;
        const wchar_t *filename;

        filename = g_default_options_filename;
        if (argc >= 3) {
            filename = argv[2];
        }

        exit_code = run_exportoptions(filename);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"importoptions")) {
        int exit_code;
        const wchar_t *filename;

        filename = g_default_options_filename;
        if (argc >= 3) {
            filename = argv[2];
        }

        exit_code = run_importoptions(filename);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"find") || command_is(command, L"list")) {
        const SCOPE *scope;
        const wchar_t *filter;
        int exit_code;

        scope = NULL;
        filter = NULL;

        if (argc >= 3) {
            scope = find_scope(argv[2]);
            if (scope != NULL) {
                if (argc >= 4) {
                    filter = argv[3];
                }
            } else {
                filter = argv[2];
            }
        }

        exit_code = run_find(scope, filter);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"remove")) {
        const SCOPE *scope;
        const wchar_t *query;
        int exit_code;

        if (argc < 4) {
            fwprintf(stderr, L"remove requires <scope> and <entry-name-or-label>.\n");
            print_usage();
            LocalFree(argv);
            return 1;
        }

        scope = find_scope(argv[2]);
        if (scope == NULL) {
            fwprintf(stderr, L"Unknown scope: %ls\n", argv[2]);
            print_scopes();
            LocalFree(argv);
            return 1;
        }

        query = argv[3];
        exit_code = run_remove(scope, query);
        LocalFree(argv);
        return exit_code;
    }

    if (command_is(command, L"removebuiltin")) {
        const BUILTIN_TARGET *target;
        int exit_code;

        if (argc < 3) {
            fwprintf(stderr, L"removebuiltin requires <name-or-label>.\n");
            print_usage();
            LocalFree(argv);
            return 1;
        }

        target = find_builtin_target(argv[2]);
        if (target == NULL) {
            fwprintf(stderr, L"Unknown built-in target: %ls\n", argv[2]);
            run_builtins(NULL);
            LocalFree(argv);
            return 1;
        }

        exit_code = run_removebuiltin(target);
        LocalFree(argv);
        return exit_code;
    }

    fwprintf(stderr, L"Unknown command: %ls\n", command);
    print_usage();
    LocalFree(argv);
    return 1;
}
