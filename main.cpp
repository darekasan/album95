#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <stdio.h>
#include <shlobj.h>

#include <vector>
#include <string>
#include <map>
#include <algorithm>

#include "bass.h"

// ---- stb_image (PNG/JPEG decode) ----
#define STBI_NO_BMP
#define STBI_NO_GIF
#define STBI_NO_PSD
#define STBI_NO_PIC
#define STBI_NO_PNM
#define STBI_NO_TGA
#define STBI_NO_HDR

#define STB_IMAGE_IMPLEMENTATION
#include "stb_image.h"

// stb_image.h が正しいかの簡易チェック
#ifndef STBI_VERSION
#error "stb_image.h is not the official stb_image.h (missing STBI_VERSION). Replace stb_image.h with the one from nothings/stb."
#endif

#define WM_APP_STARTSCAN (WM_APP + 100)
#define WM_APP_PLAYSELECT (WM_APP + 101)
#define WM_APP_NEXTTRACK (WM_APP + 102)
#define WM_APP_SCANPROGRESS (WM_APP + 103)
#define WM_APP_SCANDONE     (WM_APP + 104)

#define WM_APP_TOGGLEPLAY   (WM_APP + 10)
#define WM_APP_PLAYENTER   (WM_APP + 11)

#define IDC_BTN_PREV   1001
#define IDC_BTN_PLAY   1002
#define IDC_BTN_NEXT   1003
#define IDC_SEEKBAR    1004
#define IDC_VOLBAR  1101

#define TIMER_SCANPROG      1
#define TIMER_PLAYPOS       2

static std::string g_root = "C:\\Music";
static DWORD g_bass_rate = 44100;
static HWND g_hwnd = NULL;
static HINSTANCE g_inst = NULL;

static int g_pending_sel = -1;
static volatile LONG g_scan_count = 0;
static volatile LONG g_scan_done  = 0;
static HANDLE g_scan_thread = NULL;
static std::string g_scan_file;
static CRITICAL_SECTION g_scan_cs;

static HWND g_status = NULL;

static HWND g_tip = NULL;
static char g_tipbuf[256];

static HWND g_tracks = NULL;
static std::string g_current_album_key;

static HFONT g_ui_font = NULL;
static HBRUSH g_brush_btnface = NULL;

static WNDPROC g_old_list_proc = NULL;
static WNDPROC g_old_tracks_proc = NULL;

static WNDPROC g_old_seekbar_proc = NULL;
static WNDPROC g_old_volbar_proc = NULL;

static HWND g_btn_prev = NULL;
static HWND g_btn_play = NULL;
static HWND g_btn_next = NULL;
static HWND g_seekbar  = NULL;
static HWND g_volbar = NULL;
static HWND g_vol_label;
static int g_seek_dragging = 0;

#ifndef BASS_TAG_MP4
#define BASS_TAG_MP4 7
#endif

#ifndef TTM_SETMAXTIPWIDTH
#define TTM_SETMAXTIPWIDTH (WM_USER + 24)
#endif

#ifndef LVM_SETTOOLTIPS
#define LVM_SETTOOLTIPS (LVM_FIRST + 74)
#endif

#ifndef LVS_EX_FULLROWSELECT
#define LVS_EX_FULLROWSELECT 0x00000020
#endif

#ifndef LVM_SETEXTENDEDLISTVIEWSTYLE
#define LVM_SETEXTENDEDLISTVIEWSTYLE (LVM_FIRST + 54)
#endif

#ifndef LVM_GETEXTENDEDLISTVIEWSTYLE
#define LVM_GETEXTENDEDLISTVIEWSTYLE (LVM_FIRST + 55)
#endif

#ifndef ListView_SetExtendedListViewStyle
#define ListView_SetExtendedListViewStyle(hwndLV, dw) \
    ((DWORD)SendMessage((hwndLV), LVM_SETEXTENDEDLISTVIEWSTYLE, 0, (LPARAM)(dw)))
#endif

#define THUMB_W 96
#define THUMB_H 96

#define CONTROLBAR_HEIGHT   32
#define STATUSBAR_HEIGHT    20

static void refresh_album_list_labels();
static void populate_tracks(const std::string& album_key);

static void update_playback_status();
static void format_mmss(int sec, char* out);

static void update_play_button_label();

static const char* basenameA(const std::string& s);

static LONG WINAPI CrashFilter(EXCEPTION_POINTERS* ep) {
    char buf[256];
    wsprintfA(buf, "CRASH: code=0x%08X addr=0x%08X",
              (unsigned)ep->ExceptionRecord->ExceptionCode,
              (unsigned)ep->ExceptionRecord->ExceptionAddress);
    MessageBoxA(NULL, buf, "Album95", MB_OK | MB_ICONERROR);
    return EXCEPTION_EXECUTE_HANDLER;
}

enum StatusMode {
    STATUSMODE_IDLE = 0,
    STATUSMODE_SCANNING,
    STATUSMODE_BUILDING
};

static int g_status_mode = STATUSMODE_IDLE;

struct Meta {
    std::string album;
    std::string albumartist;
    std::string artist;
    std::string title;
    
    int track;
    int disc;
    int compilation;
    int year;
    
    Meta() : track(0), disc(0), compilation(0), year(0) {}
};

struct Track {
    std::string path;
    Meta meta;
    int duration_sec;
    Track() : duration_sec(0) {}
    Track(const std::string& p) : path(p), duration_sec(0) {}
};

struct Album {
    std::string folder;
    std::string album_name_u8;
    std::string albumartist_u8;
    std::string disp_name;
    int year;
    std::vector<Track> tracks;
    HBITMAP cover;
    Album() : year(0), cover(NULL) {}
};

static std::string album_key(const Meta& m) {
    std::string aa = !m.albumartist.empty() ? m.albumartist : m.artist;
    if (aa.empty()) aa = "Unknown Artist";
    std::string al = !m.album.empty() ? m.album : "Unknown Album";
    return aa + " - " + al;
}

static std::map<std::string, Album> g_albums;
static std::vector<std::string> g_album_order;
static HWND g_list = NULL;
static HIMAGELIST g_img = NULL;
static HSTREAM g_stream = 0;

static std::string g_now_album_key;
static int g_now_index = -1;
static HSYNC g_end_sync = 0;

static std::string album_key_from_list_index(int index)
{
    if (index < 0) return "";
    if ((size_t)index >= g_album_order.size()) return "";
    return g_album_order[index];
}

static const char* find_double_nul(const char* p, size_t maxBytes) {
    for (size_t i = 0; i + 1 < maxBytes; ++i) {
        if (p[i] == '\0' && p[i+1] == '\0') return p + i;
    }
    return NULL;
}

static size_t safe_strnlen0(const char* s, const char* end) {
    const char* p = s;
    while (p < end && *p) ++p;
    return (size_t)(p - s);
}

static std::string tolower_ascii(const std::string& s) {
    std::string o = s;
    for (size_t i=0;i<o.size();++i) {
        unsigned char c = (unsigned char)o[i];
        if (c>='A' && c<='Z') o[i] = (char)(c - 'A' + 'a');
    }
    return o;
}


static std::wstring utf8_to_wide(const std::string& s)
{
    std::wstring out;

    for (size_t i = 0; i < s.size();) {
        unsigned char c = (unsigned char)s[i];

        if (c < 0x80) {
            out.push_back((wchar_t)c);
            i++;
        }
        else if ((c & 0xE0) == 0xC0 && i + 1 < s.size()) {
            wchar_t w = (wchar_t)(((c & 0x1F) << 6) |
                                  ((unsigned char)s[i+1] & 0x3F));
            out.push_back(w);
            i += 2;
        }
        else if ((c & 0xF0) == 0xE0 && i + 2 < s.size()) {
            wchar_t w = (wchar_t)(((c & 0x0F) << 12) |
                                  (((unsigned char)s[i+1] & 0x3F) << 6) |
                                  ((unsigned char)s[i+2] & 0x3F));
            out.push_back(w);
            i += 3;
        }
        else if ((c & 0xF8) == 0xF0 && i + 3 < s.size()) {
            unsigned long cp =
                ((unsigned long)(c & 0x07) << 18) |
                ((unsigned long)((unsigned char)s[i+1] & 0x3F) << 12) |
                ((unsigned long)((unsigned char)s[i+2] & 0x3F) << 6) |
                (unsigned long)((unsigned char)s[i+3] & 0x3F);

            if (cp >= 0x10000) {
                cp -= 0x10000;
                out.push_back((wchar_t)(0xD800 + (cp >> 10)));
                out.push_back((wchar_t)(0xDC00 + (cp & 0x3FF)));
            }

            i += 4;
        }
        else {
            i++; // skip invalid
        }
    }

    return out;
}

static std::string utf8_to_acp(const std::string& u8)
{
    std::wstring w = utf8_to_wide(u8);

    int len = WideCharToMultiByte(
        CP_ACP, 0,
        w.c_str(),
        (int)w.size(),
        NULL, 0, NULL, NULL);

    if (len <= 0)
        return "";

    std::vector<char> buf(len + 1);

    WideCharToMultiByte(
        CP_ACP, 0,
        w.c_str(),
        (int)w.size(),
        &buf[0], len, NULL, NULL);

    buf[len] = 0;

    return std::string(&buf[0]);
}

static unsigned short read_be16(FILE* fp)
{
    unsigned char b[2];
    if (fread(b, 1, 2, fp) != 2) return 0;
    return (unsigned short)(((unsigned short)b[0] << 8) | b[1]);
}

static unsigned long read_be32(FILE* fp)
{
    unsigned char b[4];
    if (fread(b, 1, 4, fp) != 4) return 0;
    return ((unsigned long)b[0] << 24) |
           ((unsigned long)b[1] << 16) |
           ((unsigned long)b[2] <<  8) |
           ((unsigned long)b[3]);
}

static unsigned __int64 read_be64(FILE* fp)
{
    unsigned char b[8];
    if (fread(b, 1, 8, fp) != 8) return 0;
    return ((unsigned __int64)b[0] << 56) |
           ((unsigned __int64)b[1] << 48) |
           ((unsigned __int64)b[2] << 40) |
           ((unsigned __int64)b[3] << 32) |
           ((unsigned __int64)b[4] << 24) |
           ((unsigned __int64)b[5] << 16) |
           ((unsigned __int64)b[6] <<  8) |
           ((unsigned __int64)b[7]);
}

static int atom_type_eq(const char type[4], const char* s)
{
    return type[0] == s[0] &&
           type[1] == s[1] &&
           type[2] == s[2] &&
           type[3] == s[3];
}

static void format_mmss(int sec, char* out)
{
    if (sec < 0) sec = 0;

    int m = sec / 60;
    int s = sec % 60;

    wsprintfA(out, "%d:%02d", m, s);
}

static int read_m4a_duration_sec_from_mvhd(const char* path)
{
    FILE* fp = fopen(path, "rb");
    if (!fp) return 0;

    for (;;) {
        long atom_start = ftell(fp);
        if (atom_start < 0) break;

        unsigned long size32;
        char type[4];
        unsigned __int64 atom_size;
        long payload_start;

        size32 = read_be32(fp);
        if (size32 == 0) break;
        if (fread(type, 1, 4, fp) != 4) break;

        atom_size = size32;
        if (size32 == 1) {
            atom_size = read_be64(fp);
        }

        payload_start = ftell(fp);
        if (payload_start < 0) break;
        if (atom_size < (unsigned __int64)(payload_start - atom_start)) break;

        if (atom_type_eq(type, "moov")) {
            unsigned __int64 moov_end = atom_start + atom_size;

            while ((unsigned __int64)ftell(fp) + 8 <= moov_end) {
                long sub_start = ftell(fp);
                unsigned long sub_size32;
                char sub_type[4];
                unsigned __int64 sub_size;
                long sub_payload_start;

                sub_size32 = read_be32(fp);
                if (sub_size32 == 0) break;
                if (fread(sub_type, 1, 4, fp) != 4) break;

                sub_size = sub_size32;
                if (sub_size32 == 1) {
                    sub_size = read_be64(fp);
                }

                sub_payload_start = ftell(fp);
                if (sub_payload_start < 0) break;
                if (sub_size < (unsigned __int64)(sub_payload_start - sub_start)) break;

                if (atom_type_eq(sub_type, "mvhd")) {
                    unsigned char version;
                    unsigned char flags[3];
                    unsigned long timescale = 0;
                    double sec = 0.0;

                    if (fread(&version, 1, 1, fp) != 1) break;
                    if (fread(flags, 1, 3, fp) != 3) break;

                    if (version == 0) {
                        /* creation_time, modification_time */
                        read_be32(fp);
                        read_be32(fp);
                        timescale = read_be32(fp);
                        {
                            unsigned long duration = read_be32(fp);
                            if (timescale != 0)
                                sec = (double)duration / (double)timescale;
                        }
                    } else if (version == 1) {
                        /* creation_time, modification_time */
                        read_be64(fp);
                        read_be64(fp);
                        timescale = read_be32(fp);
                        {
                            unsigned __int64 duration = read_be64(fp);
                            if (timescale != 0)
                                sec = (double)duration / (double)timescale;
                        }
                    }

                    fclose(fp);

                    if (sec < 0.0) sec = 0.0;
                    return (int)(sec + 0.5);
                }

                if (sub_size == 0) break;
                if (fseek(fp, sub_start + (long)sub_size, SEEK_SET) != 0) break;
            }

            break;
        }

        if (atom_size == 0) break;
        if (fseek(fp, atom_start + (long)atom_size, SEEK_SET) != 0) break;
    }

    fclose(fp);
    return 0;
}

static bool mp4_get_safe(const char* tags, const char* key, std::string& out) {
    if (!tags) return false;

    // タグ領域は最大64KBだけ信用する（安全弁）
    const size_t MAX = 64 * 1024;
    const char* end0 = find_double_nul(tags, MAX);
    if (!end0) return false; // 終端が見つからない＝文字列リストではない/壊れてる

    size_t klen = strlen(key);

    const char* p = tags;
    while (p < end0 && *p) {
        size_t len = safe_strnlen0(p, end0);
        if (len == 0) break;

        if (len > klen + 1 && !strncmp(p, key, klen) && p[klen] == '=') {
            out.assign(p + klen + 1, p + len);
            return true;
        }
        p += len + 1;
    }
    return false;
}

static bool ends_with_icase(const std::string& s, const char* ext) {
    size_t n = strlen(ext);
    if (s.size() < n) return false;
    return _stricmp(s.c_str() + (s.size() - n), ext) == 0;
}

static int parse_int_prefix(const std::string& s) {
    int v = 0;
    for (size_t i=0;i<s.size();++i) {
        char c = s[i];
        if (c<'0' || c>'9') break;
        v = v*10 + (c-'0');
    }
    return v;
}

static int parse_boolish(const std::string& s) {
    if (s.empty()) return 0;
    if (s == "1" || s == "true" || s == "True" || s == "YES" || s == "yes") return 1;
    return 0;
}

static int is_digit_char(char c)
{
    return (c >= '0' && c <= '9');
}

static int parse_leading_disc_track_from_name(const std::string& name, int* out_disc, int* out_track)
{
    size_t i = 0;
    int disc = 0;
    int track = 0;

    while (i < name.size() && (name[i] == ' ' || name[i] == '\t'))
        i++;

    if (i >= name.size() || !is_digit_char(name[i]))
        return 0;

    while (i < name.size() && is_digit_char(name[i])) {
        disc = disc * 10 + (name[i] - '0');
        i++;
    }

    if (i >= name.size() || name[i] != '-')
        return 0;
    i++;

    if (i >= name.size() || !is_digit_char(name[i]))
        return 0;

    while (i < name.size() && is_digit_char(name[i])) {
        track = track * 10 + (name[i] - '0');
        i++;
    }

    if (disc <= 0 || track <= 0)
        return 0;

    if (out_disc)  *out_disc = disc;
    if (out_track) *out_track = track;
    return 1;
}

static std::string get_album_display_name(const std::string& key, Album& a)
{
    std::string base = utf8_to_acp(a.album_name_u8);
    if (key == g_current_album_key)
        return "> " + base;
    return base;
}

static std::string get_album_display_name_u8(const Album& a)
{
    if (!a.albumartist_u8.empty())
        return a.albumartist_u8 + " - " + a.album_name_u8;

    return a.album_name_u8;
}

struct AlbumDisplayNameLess
{
    bool operator()(const std::string& ka, const std::string& kb) const
    {
        std::map<std::string, Album>::const_iterator ita = g_albums.find(ka);
        std::map<std::string, Album>::const_iterator itb = g_albums.find(kb);

        if (ita == g_albums.end()) return false;
        if (itb == g_albums.end()) return true;

        const Album& a = ita->second;
        const Album& b = itb->second;

        std::string na = get_album_display_name_u8(a);
        std::string nb = get_album_display_name_u8(b);

        int cmp = _stricmp(na.c_str(), nb.c_str());
        if (cmp != 0)
            return cmp < 0;

        return _stricmp(ka.c_str(), kb.c_str()) < 0;
    }
};

static void rebuild_album_order()
{
    g_album_order.clear();

    for (std::map<std::string, Album>::iterator it = g_albums.begin();
         it != g_albums.end();
         ++it)
    {
        g_album_order.push_back(it->first);
    }

    std::sort(g_album_order.begin(), g_album_order.end(), AlbumDisplayNameLess());
}

static int ord0_last(int v) { return (v <= 0) ? 9999 : v; }

static int disc_ord(int v)  { return (v <= 0) ? 1    : v; }
static int track_ord(int v) { return (v <= 0) ? 9999 : v; }

static bool track_sort_album(const Track& a, const Track& b)
{
    int ad = disc_ord(a.meta.disc);
    int bd = disc_ord(b.meta.disc);
    if (ad != bd) return ad < bd;

    int at = track_ord(a.meta.track);
    int bt = track_ord(b.meta.track);
    if (at != bt) return at < bt;

    if (a.meta.title != b.meta.title) return a.meta.title < b.meta.title;
    return a.path < b.path;
}

static bool read_meta_m4a(const char* path, Meta& m) {
    HSTREAM ch = BASS_StreamCreateFile(FALSE, path, 0, 0, BASS_STREAM_DECODE);
    if (!ch) return false;

    const char* tags = (const char*)BASS_StreamGetTags(ch, BASS_TAG_MP4);

    if (tags) {
        mp4_get_safe(tags, "album", m.album);
        if (!mp4_get_safe(tags, "albumartist", m.albumartist)) {
            mp4_get_safe(tags, "aART", m.albumartist);
        }
        mp4_get_safe(tags, "artist", m.artist);

        if (!mp4_get_safe(tags, "title", m.title)) {
            mp4_get_safe(tags, "\xA9""nam", m.title);
        }

        std::string cpil;
        if (mp4_get_safe(tags, "cpil", cpil) || mp4_get_safe(tags, "compilation", cpil)) {
            m.compilation = parse_boolish(cpil);
        }

        std::string trk, dsk;
        if (mp4_get_safe(tags, "track", trk)) m.track = parse_int_prefix(trk);
        if (mp4_get_safe(tags, "discnumber",  dsk)) m.disc  = parse_int_prefix(dsk);

        std::string date;
        if (mp4_get_safe(tags, "year", date) || mp4_get_safe(tags, "date", date)) {
            m.year = parse_int_prefix(date);
        }
    }

    BASS_StreamFree(ch);
    return true;
}

static std::string make_album_key(const Meta& m) {
    std::string al = !m.album.empty() ? m.album : "Unknown Album";

    std::string aa;
    if (!m.albumartist.empty()) {
        aa = m.albumartist;
    } else if (m.compilation) {
        aa = "Various Artists";
    } else if (!m.artist.empty()) {
        aa = m.artist;
    } else {
        aa = "Unknown Artist";
    }

    return aa + " - " + al;
}

static std::string get_cache_dir()
{
    char exe[MAX_PATH] = {0};
    GetModuleFileNameA(NULL, exe, MAX_PATH);
    char* p = strrchr(exe, '\\');
    if (p) *p = 0;
    std::string dir = std::string(exe) + "\\cache";
    CreateDirectoryA(dir.c_str(), NULL);
    return dir;
}

static std::string get_ini_path()
{
    char exe[MAX_PATH] = {0};
    GetModuleFileNameA(NULL, exe, MAX_PATH);

    char* p = strrchr(exe, '\\');
    if (p) *p = 0;

    return std::string(exe) + "\\album95.ini";
}

static void load_settings_from_ini()
{
    char buf[MAX_PATH] = {0};
    std::string ini = get_ini_path();

    GetPrivateProfileStringA(
        "General",
        "MusicRoot",
        "C:\\Music",
        buf,
        MAX_PATH,
        ini.c_str()
    );

    if (buf[0])
        g_root = buf;
    else
        g_root = "C:\\Music";

    char ratebuf[64] = {0};
    GetPrivateProfileStringA(
        "General",
        "SampleRate",
        "44100",
        ratebuf,
        sizeof(ratebuf),
        ini.c_str()
    );

    int rate = atoi(ratebuf);
    if (rate == 32000)
        g_bass_rate = 32000;
    else
        g_bass_rate = 44100;
}

static void save_settings_to_ini()
{
    std::string ini = get_ini_path();

    WritePrivateProfileStringA(
        "General",
        "MusicRoot",
        g_root.c_str(),
        ini.c_str()
    );

    char ratebuf[32];
    wsprintfA(ratebuf, "%lu", (unsigned long)g_bass_rate);

    WritePrivateProfileStringA(
        "General",
        "SampleRate",
        ratebuf,
        ini.c_str()
    );
}

static int is_shift_pressed_at_startup()
{
    return (GetAsyncKeyState(VK_SHIFT) & 0x8000) ? 1 : 0;
}

static void choose_startup_sample_rate()
{
    int r = MessageBoxA(
        NULL,
        "Use 32 kHz output sample rate?\n\n"
        "Yes = 32000 Hz\n"
        "No  = 44100 Hz",
        "Album95",
        MB_YESNO | MB_ICONQUESTION
    );

    if (r == IDYES)
        g_bass_rate = 32000;
    else
        g_bass_rate = 44100;
}

static unsigned int fnv1a32(const char* s)
{
    unsigned int h = 2166136261u;
    while (*s) {
        h ^= (unsigned char)*s++;
        h *= 16777619u;
    }
    return h;
}

static std::string make_cache_path_for_key(const std::string& key)
{
    char buf[64];
    unsigned int h = fnv1a32(key.c_str());
    wsprintfA(buf, "%08X.bmp", h);
    return get_cache_dir() + "\\" + buf;
}

static HBITMAP load_cached_bmp(const char* path)
{
    return (HBITMAP)LoadImageA(
        NULL, path, IMAGE_BITMAP, 0, 0,
        LR_LOADFROMFILE | LR_CREATEDIBSECTION
    );
}

#pragma pack(push,1)
struct BMPFILEHDR {
    unsigned short bfType;      // 'BM'
    unsigned int   bfSize;
    unsigned short bfReserved1;
    unsigned short bfReserved2;
    unsigned int   bfOffBits;
};
#pragma pack(pop)

static bool save_hbitmap_as_bmp24(const char* path, HBITMAP hbmp, int w, int h)
{
    if (!hbmp || w<=0 || h<=0) return false;

    HDC hdc = GetDC(NULL);

    BITMAPINFO bi;
    ZeroMemory(&bi, sizeof(bi));
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = w;
    bi.bmiHeader.biHeight = h;          // BMPはbottom-upで保存するのが簡単
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 24;
    bi.bmiHeader.biCompression = BI_RGB;

    // 24bppは1行が4byte境界
    int stride = ((w * 3 + 3) / 4) * 4;
    int datasz = stride * h;

    std::vector<unsigned char> pixels(datasz);

    // GetDIBitsで変換コピー（16bpp/32bppなど何でも24bppにしてくれる）
    int got = GetDIBits(hdc, hbmp, 0, (UINT)h, &pixels[0], &bi, DIB_RGB_COLORS);

    ReleaseDC(NULL, hdc);

    if (got != h) return false;

    BMPFILEHDR fh;
    fh.bfType = 0x4D42; // 'BM'
    fh.bfOffBits = sizeof(BMPFILEHDR) + sizeof(BITMAPINFOHEADER);
    fh.bfSize = fh.bfOffBits + datasz;
    fh.bfReserved1 = 0;
    fh.bfReserved2 = 0;

    FILE* f = fopen(path, "wb");
    if (!f) return false;

    fwrite(&fh, 1, sizeof(fh), f);
    fwrite(&bi.bmiHeader, 1, sizeof(BITMAPINFOHEADER), f);
    fwrite(&pixels[0], 1, datasz, f);
    fclose(f);
    return true;
}

static void add_track(const std::string& folder, const std::string& file, const Meta& m) {
    Album& a = g_albums[folder];
    a.folder = folder;
    Track t(file);
    t.meta = m;
    a.tracks.push_back(t);
}

static void add_track2(const std::string& key, const std::string& file, const Meta& m, int duration_sec) {
    Album& a = g_albums[key];
    a.folder = key;

    if (a.album_name_u8.empty()) {
        a.album_name_u8 = !m.album.empty() ? m.album : "Unknown Album";
    }

    Track t(file);
    t.meta = m;
    t.duration_sec = duration_sec;
    a.tracks.push_back(t);

    if (a.year == 0 && m.year > 0) a.year = m.year;
}

static int browse_for_music_folder(HWND owner, std::string& out_path)
{
    char display[MAX_PATH] = {0};

    BROWSEINFOA bi;
    ZeroMemory(&bi, sizeof(bi));
    bi.hwndOwner = owner;
    bi.pszDisplayName = display;
    bi.lpszTitle = "Select music folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS;

    LPITEMIDLIST pidl = SHBrowseForFolderA(&bi);
    if (!pidl)
        return 0;

    char path[MAX_PATH] = {0};
    int ok = SHGetPathFromIDListA(pidl, path);

    LPMALLOC pMalloc = NULL;
    if (SUCCEEDED(SHGetMalloc(&pMalloc)) && pMalloc) {
        pMalloc->Free(pidl);
        pMalloc->Release();
    }

    if (!ok || !path[0])
        return 0;

    out_path = path;
    return 1;
}

static void scan_folder(const std::string& path) {
    WIN32_FIND_DATAA fd;
    std::string pattern = path + "\\*";
    HANDLE h = FindFirstFileA(pattern.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return;

    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            if (strcmp(fd.cFileName, ".") != 0 && strcmp(fd.cFileName, "..") != 0) {
                scan_folder(path + "\\" + fd.cFileName);
            }
        } else {
            std::string file = path + "\\" + fd.cFileName;
            if (ends_with_icase(file, ".m4a")) {
                Meta m;
                int duration_sec = 0;

                read_meta_m4a(file.c_str(), m);

                if (m.disc <= 0 || m.track <= 0) {
                    int fd_disc = 0;
                    int fd_track = 0;

                    if (parse_leading_disc_track_from_name(fd.cFileName, &fd_disc, &fd_track)) {
                        if (m.disc <= 0)
                            m.disc = fd_disc;
                        if (m.track <= 0)
                            m.track = fd_track;
                    }
                }

                std::string key = make_album_key(m);
                duration_sec = read_m4a_duration_sec_from_mvhd(file.c_str());
                add_track2(key, file, m, duration_sec);

                InterlockedIncrement(&g_scan_count);

                EnterCriticalSection(&g_scan_cs);
                g_scan_file = file;
                LeaveCriticalSection(&g_scan_cs);
            }
        }
    } while (FindNextFileA(h, &fd));

    FindClose(h);
}

static DWORD WINAPI ScanThreadProc(LPVOID) {
    g_scan_count = 0;
    g_scan_done  = 0;

    scan_folder(g_root);

    g_scan_done = 1;
    if (g_hwnd) PostMessageA(g_hwnd, WM_APP_SCANDONE, 0, 0);
    return 0;
}

static void sort_all_albums()
{
    for (std::map<std::string, Album>::iterator it = g_albums.begin();
         it != g_albums.end(); ++it) {
        Album& a = it->second;
        std::sort(a.tracks.begin(), a.tracks.end(), track_sort_album);
    }
}

static void regroup_by_album_with_va_fallback() {
    // 1) いまのg_albumsから全Trackを平坦化
    std::vector<Track> all;
    for (std::map<std::string, Album>::iterator it = g_albums.begin(); it != g_albums.end(); ++it) {
        Album& a = it->second;
        for (size_t i=0;i<a.tracks.size();++i) all.push_back(a.tracks[i]);
    }

    // 2) album名ごとに artist種類数を数える
    std::map<std::string, std::map<std::string, int> > album_artists; // album -> (artist -> count)
    for (size_t i=0;i<all.size();++i) {
        const Meta& m = all[i].meta;
        std::string al = !m.album.empty() ? m.album : "Unknown Album";
        std::string ar = !m.artist.empty() ? m.artist : "";
        album_artists[al][ar] += 1;
    }

    // 3) g_albumsを作り直す
    g_albums.clear();

    for (size_t i=0;i<all.size();++i) {
        Track t = all[i];
        Meta m = t.meta;

        std::string al = !m.album.empty() ? m.album : "Unknown Album";

        // ★ albumartistが空で、同アルバムに複数artistがいるならVA扱い
        if (m.albumartist.empty()) {
            int kinds = 0;
            std::map<std::string,int>& mp = album_artists[al];
            for (std::map<std::string,int>::iterator it = mp.begin(); it != mp.end(); ++it) {
                if (!it->first.empty()) ++kinds;
            }
            if (kinds >= 2) {
                m.albumartist = "Various Artists";
            }
        }

        std::string key = make_album_key(m);

        Album& a = g_albums[key];
        a.folder = key;

        if (a.album_name_u8.empty())
            a.album_name_u8 = !m.album.empty() ? m.album : "Unknown Album";
            
        if (a.albumartist_u8.empty())
            a.albumartist_u8 = m.albumartist;

        t.meta = m;

        a.tracks.push_back(t);
    }
}

static const char* basenameA(const std::string& s){
    size_t p = s.find_last_of("\\/");
    return (p==std::string::npos) ? s.c_str() : (s.c_str()+p+1);
}

// ---- MP4 covr 抽出（簡易） ----
// 厳密なMP4パーサではないが、多くのm4aのcovr抽出に通る “雑実装”。
// まず動作優先。後でちゃんとboxツリーを辿る版に置換可能。

// buf はファイルの [base_off .. base_off+buf.size()) を読んだもの
// 見つかったら img_off/img_size を「ファイル先頭からの絶対オフセット」で返す
static unsigned int u32be_at(const unsigned char* p) {
    return ((unsigned)p[0] << 24) | ((unsigned)p[1] << 16) | ((unsigned)p[2] << 8) | (unsigned)p[3];
}

// buf はファイルの [base_off .. base_off+buf.size()) を読んだもの
// 見つかったら img_off/img_size を「ファイル先頭からの絶対オフセット」で返す
static bool scan_covr_find_data(const std::vector<unsigned char>& buf,
                               long base_off,
                               long file_size,
                               long& img_off,
                               long& img_size)
{
    const long n = (long)buf.size();
    if (n < 32) return false;

    for (long i = 0; i <= n - 8; ++i) {
        if (memcmp(&buf[i], "covr", 4) != 0) continue;

        long j_end = i + 256 * 1024;
        if (j_end > n - 8) j_end = n - 8;

        for (long j = i; j < j_end; ++j) {
            if (memcmp(&buf[j], "data", 4) != 0) continue;
            if (j < 4) continue;

            // data atom: [size(4)][type(4)="data"][flags+type(8)][payload...]
            const long atom_start_in_buf = j - 4;

            // ヘッダ16バイトだけは buf 内に必要
            const long header = 16;
            if (atom_start_in_buf + header > n) continue;

            unsigned int atom_size_be = u32be_at(&buf[atom_start_in_buf]);
            long atom_size = (long)atom_size_be;

            // sanity
            if (atom_size < header) continue;
            // ファイル上の atom_start の絶対位置
            long atom_start_abs = base_off + atom_start_in_buf;
            if (atom_start_abs < 0 || atom_start_abs >= file_size) continue;

            long payload_abs = atom_start_abs + header;
            long payload_size = atom_size - header;

            // ファイル範囲チェック
            if (payload_abs < 0) continue;
            if (payload_abs + payload_size > file_size) continue;

            img_off = payload_abs;
            img_size = payload_size;
            return true;
        }
    }
    return false;
}

static bool extract_covr(const char* path, std::vector<unsigned char>& out)
{
    out.clear();

    FILE* f = fopen(path, "rb");
    if (!f) return false;

    fseek(f, 0, SEEK_END);
    long fsize = ftell(f);
    if (fsize <= 0) { fclose(f); return false; }

    const long SCAN = 512 * 1024;                 // 探索窓（軽い）
    const long MAX_IMG = 8 * 1024 * 1024;         // ★画像は5MBまで許可

    std::vector<unsigned char> buf;
    buf.resize((fsize < SCAN) ? fsize : SCAN);

    long img_off = -1, img_size = 0;

    // ---- 先頭を読む ----
    fseek(f, 0, SEEK_SET);
    if (fread(&buf[0], 1, buf.size(), f) == buf.size()) {
        if (scan_covr_find_data(buf, 0, fsize, img_off, img_size)) {
            // found
        }
    }

    // ---- 見つからなければ末尾 ----
    if (img_off < 0 && fsize > (long)buf.size()) {
        long base = fsize - (long)buf.size();
        fseek(f, base, SEEK_SET);
        if (fread(&buf[0], 1, buf.size(), f) == buf.size()) {
            scan_covr_find_data(buf, base, fsize, img_off, img_size);
        }
    }

    if (img_off < 0 || img_size <= 0) { fclose(f); return false; }
    if (img_size > MAX_IMG) { fclose(f); return false; }   // ★5MB超えは拒否

    // ★ここがポイント：payloadだけ別途読む
    out.resize((size_t)img_size);
    fseek(f, img_off, SEEK_SET);
    size_t got = fread(&out[0], 1, (size_t)img_size, f);
    fclose(f);

    if (got != (size_t)img_size) {
        out.clear();
        return false;
    }
    return true;
}

static bool looks_like_png_or_jpeg(const std::vector<unsigned char>& v){
    if (v.size() >= 8 && !memcmp(&v[0], "\x89PNG\r\n\x1a\n", 8)) return true;
    if (v.size() >= 2 && v[0]==0xFF && v[1]==0xD8) return true; // JPEG
    return false;
}

static bool is_jpeg(const unsigned char* d, int size)
{
    return size > 2 && d[0] == 0xFF && d[1] == 0xD8;
}

static unsigned char* load_image_fast(const unsigned char* data, int size,
                                       int* w, int* h)
{
    int n;

    if (is_jpeg(data, size))
    {
        int iw, ih;
        if (stbi_info_from_memory(data, size, &iw, &ih, &n))
        {
            // 128px以下なら普通にロード
            if (iw <= 512 && ih <= 512)
            {
                return stbi_load_from_memory(data, size, w, h, &n, 3);
            }
        }

        // ★ JPEGはRGBのみでロード（軽い）
        return stbi_load_from_memory(data, size, w, h, &n, 3);
    }

    // PNGなど
    return stbi_load_from_memory(data, size, w, h, &n, 4);
}

static HBITMAP make_thumb_96_rgb565_from_memory(const unsigned char* data, int size)
{
    int w=0,h=0,n=0;
    unsigned char* src = stbi_load_from_memory((const stbi_uc*)data, size, &w, &h, &n, 3);
    if (!src || w<=0 || h<=0) return NULL;

    const int TW=96, TH=96;

    // 565 DIB
    struct { BITMAPINFOHEADER h; DWORD mask[3]; } bmi;
    ZeroMemory(&bmi, sizeof(bmi));
    bmi.h.biSize = sizeof(BITMAPINFOHEADER);
    bmi.h.biWidth = TW;
    bmi.h.biHeight = -TH;
    bmi.h.biPlanes = 1;
    bmi.h.biBitCount = 16;
    bmi.h.biCompression = BI_BITFIELDS;
    bmi.mask[0]=0xF800; bmi.mask[1]=0x07E0; bmi.mask[2]=0x001F;

    void* bits=NULL;
    HBITMAP bmp = CreateDIBSection(NULL, (BITMAPINFO*)&bmi, DIB_RGB_COLORS, &bits, NULL, 0);
    if (!bmp || !bits) { stbi_image_free(src); return NULL; }

    unsigned short* dst = (unsigned short*)bits;

    // 中央クロップ（正方形）
    int side = (w < h) ? w : h;
    int ox = (w - side) >> 1;
    int oy = (h - side) >> 1;

    // 固定小数(16.16)で bilinear
    // 出力(0..TW-1) -> 入力(0..side-1)
    const int FP = 16;
    const int ONE = 1 << FP;

    // 端のサンプルがはみ出ないように (side-2) までに抑える
    // ただし sideが1のときは詰むのでガード
    if (side < 2) {
        stbi_image_free(src);
        DeleteObject(bmp);
        return NULL;
    }

    // 入力座標を「ピクセル中心」寄せにする（少しだけシャープに）
    // u = (x+0.5)*side/TW - 0.5
    // を固定小数で
    const int stepX = (side << FP) / TW;
    const int stepY = (side << FP) / TH;

    int yfp = (stepY >> 1) - (ONE >> 1);
    for (int y=0; y<TH; ++y, yfp += stepY)
    {
        int sy = yfp >> FP;
        int fy = yfp & (ONE-1);
        if (sy < 0) { sy = 0; fy = 0; }
        if (sy > side-2) { sy = side-2; fy = ONE-1; }

        const unsigned char* row0 = src + ((oy + sy) * w + ox) * 3;
        const unsigned char* row1 = src + ((oy + sy + 1) * w + ox) * 3;

        unsigned short* out = dst + y*TW;

        int xfp = (stepX >> 1) - (ONE >> 1);
        for (int x=0; x<TW; ++x, xfp += stepX)
        {
            int sx = xfp >> FP;
            int fx = xfp & (ONE-1);
            if (sx < 0) { sx = 0; fx = 0; }
            if (sx > side-2) { sx = side-2; fx = ONE-1; }

            const unsigned char* p00 = row0 + sx*3;
            const unsigned char* p10 = row0 + (sx+1)*3;
            const unsigned char* p01 = row1 + sx*3;
            const unsigned char* p11 = row1 + (sx+1)*3;

            // 重み（0..65535）
            unsigned int wx1 = fx;
            unsigned int wx0 = ONE - wx1;
            unsigned int wy1 = fy;
            unsigned int wy0 = ONE - wy1;

            // 4点補間： (p00*wx0 + p10*wx1)*wy0 + (p01*wx0 + p11*wx1)*wy1
            // 途中で >>16 せず 64bitで保持して最後に >>32
            unsigned int r00=p00[0], g00=p00[1], b00=p00[2];
            unsigned int r10=p10[0], g10=p10[1], b10=p10[2];
            unsigned int r01=p01[0], g01=p01[1], b01=p01[2];
            unsigned int r11=p11[0], g11=p11[1], b11=p11[2];

            unsigned long long r0 = (unsigned long long)r00*wx0 + (unsigned long long)r10*wx1;
            unsigned long long g0 = (unsigned long long)g00*wx0 + (unsigned long long)g10*wx1;
            unsigned long long b0 = (unsigned long long)b00*wx0 + (unsigned long long)b10*wx1;

            unsigned long long r1 = (unsigned long long)r01*wx0 + (unsigned long long)r11*wx1;
            unsigned long long g1 = (unsigned long long)g01*wx0 + (unsigned long long)g11*wx1;
            unsigned long long b1 = (unsigned long long)b01*wx0 + (unsigned long long)b11*wx1;

            unsigned long long r = r0*wy0 + r1*wy1;
            unsigned long long g = g0*wy0 + g1*wy1;
            unsigned long long b = b0*wy0 + b1*wy1;

            // r/g/b は 8bit * 16bit * 16bit 相当なので >>32 で 8bitに戻る
            unsigned int rr = (unsigned int)(r >> 32);
            unsigned int gg = (unsigned int)(g >> 32);
            unsigned int bb = (unsigned int)(b >> 32);

            out[x] = (unsigned short)(((rr >> 3) << 11) | ((gg >> 2) << 5) | (bb >> 3));
        }
    }

    stbi_image_free(src);
    return bmp;
}

static void ensure_cover_cached(Album& a, const std::string& album_key)
{
    if (a.cover) return;

    std::string bmpPath = make_cache_path_for_key(album_key);

    // 1) まずキャッシュ読む
    HBITMAP cached = load_cached_bmp(bmpPath.c_str());
    if (cached) {
        a.cover = cached;
        return;
    }

    // 2) 無ければ今まで通り covr→decode
    if (!a.tracks.empty()) {
        std::vector<unsigned char> img;
        if (extract_covr(a.tracks[0].path.c_str(), img) &&
            !img.empty() &&
            img.size() <= 8 * 1024 * 1024 &&
            looks_like_png_or_jpeg(img))
        {
            a.cover = make_thumb_96_rgb565_from_memory(&img[0], (int)img.size());
            if (a.cover) {
                // 3) 保存（次回から爆速）
                save_hbitmap_as_bmp24(bmpPath.c_str(), a.cover, THUMB_W, THUMB_H);
            }
        }
    }
}

static HBITMAP image_from_memory(const unsigned char* data, int size) {
    int w = 0, h = 0, n = 0;
    unsigned char* img = stbi_load_from_memory((const stbi_uc*)data, size, &w, &h, &n, 4);
    if (!img || w <= 0 || h <= 0) return NULL;

    BITMAPINFO bmi;
    ZeroMemory(&bmi, sizeof(bmi));
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = w;
    bmi.bmiHeader.biHeight = -h;  // top-down
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = NULL;
    HBITMAP bmp = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, &bits, NULL, 0);
    if (bmp && bits) {
        memcpy(bits, img, (size_t)w * (size_t)h * 4);
    }
    stbi_image_free(img);
    return bmp;
}

static void CALLBACK on_stream_end(HSYNC /*handle*/, DWORD /*channel*/, DWORD /*data*/, DWORD /*user*/)
{
    if (g_hwnd) PostMessageA(g_hwnd, WM_APP_NEXTTRACK, 0, 0);
}

static LRESULT CALLBACK control_subclass_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_KEYDOWN) {
        if (wParam == VK_SPACE) {
            PostMessageA(g_hwnd, WM_APP_TOGGLEPLAY, 0, 0);
            return 0;
        }

        if (wParam == VK_RETURN) {
            if (hwnd == g_list) {
                int sel = ListView_GetNextItem(g_list, -1, LVNI_SELECTED);
                if (sel >= 0) {
                    g_pending_sel = sel;
                    PostMessageA(g_hwnd, WM_APP_PLAYSELECT, 0, 0);
                }
                return 0;
            }

            if (hwnd == g_tracks) {
                int sel = ListView_GetNextItem(g_tracks, -1, LVNI_SELECTED);
                if (sel >= 0 && !g_current_album_key.empty()) {
                    PostMessageA(g_hwnd, WM_APP_PLAYENTER, (WPARAM)sel, 0);
                }
                return 0;
            }
        }
    }

    if (hwnd == g_list && g_old_list_proc)
        return CallWindowProcA(g_old_list_proc, hwnd, msg, wParam, lParam);

    if (hwnd == g_tracks && g_old_tracks_proc)
        return CallWindowProcA(g_old_tracks_proc, hwnd, msg, wParam, lParam);

    if (hwnd == g_seekbar && g_old_seekbar_proc)
        return CallWindowProcA(g_old_seekbar_proc, hwnd, msg, wParam, lParam);

    if (hwnd == g_volbar && g_old_volbar_proc)
        return CallWindowProcA(g_old_volbar_proc, hwnd, msg, wParam, lParam);

    return DefWindowProcA(hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK listview_subclass_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_KEYDOWN) {
        if (wParam == VK_RETURN) {
            if (hwnd == g_list) {
                int sel = ListView_GetNextItem(g_list, -1, LVNI_SELECTED);
                if (sel >= 0) {
                    g_pending_sel = sel;
                    PostMessageA(g_hwnd, WM_APP_PLAYSELECT, 0, 0);
                }
                return 0;
            }

            if (hwnd == g_tracks) {
                int sel = ListView_GetNextItem(g_tracks, -1, LVNI_SELECTED);
                if (sel >= 0 && !g_current_album_key.empty()) {
                    PostMessageA(g_hwnd, WM_APP_PLAYENTER, (WPARAM)sel, 0);
                }
                return 0;
            }
        }
    }

    if (hwnd == g_list && g_old_list_proc)
        return CallWindowProcA(g_old_list_proc, hwnd, msg, wParam, lParam);

    if (hwnd == g_tracks && g_old_tracks_proc)
        return CallWindowProcA(g_old_tracks_proc, hwnd, msg, wParam, lParam);

    return DefWindowProcA(hwnd, msg, wParam, lParam);
}

// ---- BASS再生 ----
static void dbg(const char* s) { OutputDebugStringA(s); OutputDebugStringA("\n"); }

static void play_file(const char* path) {
    if (g_stream) {
        BASS_StreamFree(g_stream);
        g_stream = 0;
        g_end_sync = 0;
    }

    g_stream = BASS_StreamCreateFile(FALSE, path, 0, 0, 0);
    if (!g_stream) {
        char buf[256];
        wsprintfA(buf, "Open failed err=%d", BASS_ErrorGetCode());
        dbg(buf);
        return;
    }

    g_end_sync = BASS_ChannelSetSync(g_stream, BASS_SYNC_END, 0, &on_stream_end, 0);

    if (!BASS_ChannelPlay(g_stream, FALSE)) {
        char buf[128];
        wsprintfA(buf, "Play failed err=%d", BASS_ErrorGetCode());
        dbg(buf);
        return;
    }
}

static void play_track_in_album(const std::string& album_key, int index)
{
    std::map<std::string, Album>::iterator it = g_albums.find(album_key);
    if (it == g_albums.end()) return;

    Album& a = it->second;
    if (index < 0 || index >= (int)a.tracks.size()) return;

    std::string prev_album_key = g_now_album_key;
    int prev_index = g_now_index;
    int album_changed = (prev_album_key != album_key);

    g_now_album_key = album_key;
    g_now_index = index;

    g_current_album_key = album_key;

    if (album_changed) {
        refresh_album_list_labels();
        populate_tracks(album_key);
    } else {
        if (prev_index >= 0 && prev_index < ListView_GetItemCount(g_tracks))
            ListView_SetItemText(g_tracks, prev_index, 0, (LPSTR)"");
    }

    play_file(a.tracks[index].path.c_str());

    ListView_SetItemState(
        g_tracks,
        index,
        LVIS_SELECTED | LVIS_FOCUSED,
        LVIS_SELECTED | LVIS_FOCUSED
    );
    ListView_EnsureVisible(g_tracks, index, FALSE);

    ListView_SetItemText(g_tracks, index, 0, (LPSTR)">");

    update_playback_status();
    update_play_button_label();
}

static void play_next_in_album()
{
    if (g_now_album_key.empty()) return;

    std::map<std::string, Album>::iterator it = g_albums.find(g_now_album_key);
    if (it == g_albums.end()) return;

    Album& a = it->second;
    int next = g_now_index + 1;

    if (next >= (int)a.tracks.size()) {
        g_now_index = -1;
        g_now_album_key.clear();
        g_current_album_key.clear();

        if (g_stream) {
            BASS_StreamFree(g_stream);
            g_stream = 0;
            g_end_sync = 0;
        }

        refresh_album_list_labels();
        update_playback_status();
        update_play_button_label();

        if (g_seekbar)
            SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);

        return;
    }

    play_track_in_album(g_now_album_key, next);
}

static void play_prev_in_album()
{
    if (g_now_album_key.empty())
        return;

    std::map<std::string, Album>::iterator it = g_albums.find(g_now_album_key);
    if (it == g_albums.end())
        return;

    Album& a = it->second;
    if (a.tracks.empty())
        return;

    if (g_now_index < 0) {
        play_track_in_album(g_now_album_key, 0);
        return;
    }

    int prev = g_now_index - 1;
    if (prev < 0)
        prev = 0;

    play_track_in_album(g_now_album_key, prev);
}

static void play_default_album()
{
    if (g_albums.empty())
        return;

    int sel = -1;

    if (g_list)
        sel = ListView_GetNextItem(g_list, -1, LVNI_SELECTED);

    if (sel >= 0) {
        std::string key = album_key_from_list_index(sel);
        if (!key.empty()) {
            std::map<std::string, Album>::iterator it = g_albums.find(key);
            if (it != g_albums.end() && !it->second.tracks.empty()) {
                play_track_in_album(key, 0);
                return;
            }
        }
    }

    if (!g_album_order.empty()) {
        const std::string& key = g_album_order[0];
        std::map<std::string, Album>::iterator it = g_albums.find(key);
        if (it != g_albums.end() && !it->second.tracks.empty()) {
            play_track_in_album(key, 0);
        }
    }
}

// ---- UI ----
static void populate_list()
{
    ListView_DeleteAllItems(g_list);
    ImageList_RemoveAll(g_img);

    int index = 0;
    int total = (int)g_album_order.size();
    int done = 0;

    for (size_t i = 0; i < g_album_order.size(); ++i)
    {
        const std::string& album_key = g_album_order[i];
        std::map<std::string, Album>::iterator it = g_albums.find(album_key);
        if (it == g_albums.end())
            continue;

        Album& a = it->second;

        ensure_cover_cached(a, album_key);

        int imgIndex = -1;
        if (a.cover)
            imgIndex = ImageList_Add(g_img, a.cover, NULL);

        LVITEMA item;
        ZeroMemory(&item, sizeof(item));
        item.mask = LVIF_TEXT | LVIF_IMAGE;
        item.iItem = index++;

        a.disp_name = get_album_display_name(album_key, a);
        item.pszText = (LPSTR)a.disp_name.c_str();
        item.iImage = imgIndex;

        ListView_InsertItem(g_list, &item);

        done++;

        char buf[256];
        wsprintfA(buf, "Building album list %d / %d", done, total);
        if (g_status)
            SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)buf);

        MSG msg;
        while (PeekMessageA(&msg, NULL, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessageA(&msg);
        }
    }
}

static void toggle_play_pause()
{
    if (!g_stream) return;

    DWORD st = BASS_ChannelIsActive(g_stream);
    if (st == BASS_ACTIVE_PLAYING) {
        BASS_ChannelPause(g_stream);
    } else if (st == BASS_ACTIVE_PAUSED) {
        BASS_ChannelPlay(g_stream, FALSE);
    }
}

static void populate_tracks(const std::string& album_key)
{
    if (!g_tracks) return;

    ListView_DeleteAllItems(g_tracks);

    std::map<std::string, Album>::iterator it = g_albums.find(album_key);
    if (it == g_albums.end()) return;

    Album& a = it->second;

    int index = 0;
    for (size_t i = 0; i < a.tracks.size(); ++i) {
        Track& t = a.tracks[i];

        LVITEMA item;
        ZeroMemory(&item, sizeof(item));
        item.mask = LVIF_TEXT;
        item.iItem = index;
        item.iSubItem = 0;
        item.pszText = (LPSTR)"";
        ListView_InsertItem(g_tracks, &item);

        char numbuf[32];
        if (t.meta.track > 0)
            wsprintfA(numbuf, "%d", t.meta.track);
        else
            wsprintfA(numbuf, "%d", index + 1);

        std::string title_u8 = !t.meta.title.empty() ? t.meta.title : basenameA(t.path);
        std::string title_a = utf8_to_acp(title_u8);

        std::string artist_u8 = !t.meta.artist.empty() ? t.meta.artist : "";
        std::string artist_a = utf8_to_acp(artist_u8);

        char timebuf[16];
        format_mmss(t.duration_sec, timebuf);

        ListView_SetItemText(g_tracks, index, 1, numbuf);
        ListView_SetItemText(g_tracks, index, 2, (LPSTR)title_a.c_str());
        ListView_SetItemText(g_tracks, index, 3, (LPSTR)artist_a.c_str());
        ListView_SetItemText(g_tracks, index, 4, timebuf);

        ++index;
    }

    if (album_key == g_now_album_key &&
        g_now_index >= 0 &&
        g_now_index < ListView_GetItemCount(g_tracks))
    {
        ListView_SetItemText(g_tracks, g_now_index, 0, ">");
        ListView_SetItemState(
            g_tracks,
            g_now_index,
            LVIS_SELECTED | LVIS_FOCUSED,
            LVIS_SELECTED | LVIS_FOCUSED
        );
        ListView_EnsureVisible(g_tracks, g_now_index, FALSE);
    }
}

static void refresh_album_list_labels()
{
    if (!g_list) return;

    for (size_t i = 0; i < g_album_order.size(); ++i)
    {
        const std::string& key = g_album_order[i];
        std::map<std::string, Album>::iterator it = g_albums.find(key);
        if (it == g_albums.end())
            continue;

        Album& a = it->second;
        a.disp_name = get_album_display_name(key, a);
        ListView_SetItemText(g_list, (int)i, 0, (LPSTR)a.disp_name.c_str());
    }
}

static void seek_to_bar_pos(int bar_pos)
{
    if (!g_stream || g_now_album_key.empty() || g_now_index < 0)
        return;

    std::map<std::string, Album>::iterator it = g_albums.find(g_now_album_key);
    if (it == g_albums.end())
        return;

    Album& a = it->second;
    if (g_now_index < 0 || g_now_index >= (int)a.tracks.size())
        return;

    Track& t = a.tracks[g_now_index];
    int len_sec = t.duration_sec;

    if (len_sec <= 0)
        return;

    if (bar_pos < 0) bar_pos = 0;
    if (bar_pos > 1000) bar_pos = 1000;

    double sec = ((double)bar_pos * (double)len_sec) / 1000.0;
    DWORD byte_pos = BASS_ChannelSeconds2Bytes(g_stream, sec);

    BASS_ChannelSetPosition(g_stream, byte_pos);
}

static void update_seekbar_pos()
{
    if (!g_seekbar) return;
    if (g_seek_dragging) return;
    if (!g_stream || g_now_album_key.empty() || g_now_index < 0) {
        SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);
        return;
    }

    std::map<std::string, Album>::iterator it = g_albums.find(g_now_album_key);
    if (it == g_albums.end()) {
        SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);
        return;
    }

    Album& a = it->second;
    if (g_now_index < 0 || g_now_index >= (int)a.tracks.size()) {
        SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);
        return;
    }

    Track& t = a.tracks[g_now_index];
    int len_sec = t.duration_sec;
    int pos_sec = 0;

    DWORD st = BASS_ACTIVE_STOPPED;
    if (g_stream)
        st = BASS_ChannelIsActive(g_stream);

    if (g_stream && (st == BASS_ACTIVE_PLAYING || st == BASS_ACTIVE_PAUSED)) {
        DWORD pos = BASS_ChannelGetPosition(g_stream);
        if (pos != (DWORD)-1) {
            pos_sec = (int)BASS_ChannelBytes2Seconds(g_stream, pos);
        }
    }

    if (len_sec <= 0) {
        SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);
        return;
    }

    if (pos_sec < 0) pos_sec = 0;
    if (pos_sec > len_sec) pos_sec = len_sec;

    int bar_pos = (pos_sec * 1000) / len_sec;
    if (bar_pos < 0) bar_pos = 0;
    if (bar_pos > 1000) bar_pos = 1000;

    SendMessageA(g_seekbar, TBM_SETPOS, TRUE, bar_pos);
}

static void update_playback_status()
{
    if (!g_status) return;

    if (g_status_mode != STATUSMODE_IDLE)
        return;

    if (!g_stream || g_now_album_key.empty() || g_now_index < 0) {
        SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)"Ready");
        if (g_seekbar)
            SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);
        return;
    }

    std::map<std::string, Album>::iterator it = g_albums.find(g_now_album_key);
    if (it == g_albums.end()) {
        SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)"Ready");
        return;
    }

    Album& a = it->second;
    if (g_now_index < 0 || g_now_index >= (int)a.tracks.size()) {
        SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)"Ready");
        return;
    }

    Track& t = a.tracks[g_now_index];

    std::string title_u8 = !t.meta.title.empty() ? t.meta.title : basenameA(t.path);
    std::string artist_u8 = !t.meta.artist.empty() ? t.meta.artist : "";

    std::string title_a = utf8_to_acp(title_u8);
    std::string artist_a = utf8_to_acp(artist_u8);

    int pos_sec = 0;
    int len_sec = t.duration_sec;

    DWORD st = BASS_ACTIVE_STOPPED;
    if (g_stream)
        st = BASS_ChannelIsActive(g_stream);

    if (g_stream && (st == BASS_ACTIVE_PLAYING || st == BASS_ACTIVE_PAUSED)) {
        DWORD pos = BASS_ChannelGetPosition(g_stream);
        if (pos != (DWORD)-1) {
            pos_sec = (int)BASS_ChannelBytes2Seconds(g_stream, pos);
        }
    }

    if (len_sec > 0 && pos_sec > len_sec)
        pos_sec = len_sec;

    char posbuf[16];
    char lenbuf[16];
    char text[1024];

    format_mmss(pos_sec, posbuf);
    format_mmss(len_sec, lenbuf);

    const char* state = "";
    if (st == BASS_ACTIVE_PAUSED)
        state = "[Paused] ";

    if (!artist_a.empty())
        wsprintfA(text, "%s%s / %s  -  %s / %s",
                  state, posbuf, lenbuf, title_a.c_str(), artist_a.c_str());
    else
        wsprintfA(text, "%s%s / %s  -  %s",
                  state, posbuf, lenbuf, title_a.c_str());

    SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)text);
    update_seekbar_pos();
}

static int can_toggle_current_stream()
{
    if (!g_stream) return 0;

    DWORD st = BASS_ChannelIsActive(g_stream);
    return (st == BASS_ACTIVE_PLAYING || st == BASS_ACTIVE_PAUSED);
}

static void update_play_button_label()
{
    if (!g_btn_play)
        return;

    const char* text = "Play";

    if (g_stream) {
        DWORD st = BASS_ChannelIsActive(g_stream);
        if (st == BASS_ACTIVE_PLAYING)
            text = "Pause";
        else
            text = "Play";
    }

    SetWindowTextA(g_btn_play, text);
}

static void layout_children(HWND h)
{
    RECT rc;
    GetClientRect(h, &rc);

    int w = rc.right - rc.left;
    int hgt = rc.bottom - rc.top;

    int status_h = STATUSBAR_HEIGHT;
    int ctrl_h   = CONTROLBAR_HEIGHT;

    int status_top = hgt - status_h;
    int ctrl_top   = status_top - ctrl_h;
    int content_h  = ctrl_top;

    int right_w = 345;
    if (right_w > w) right_w = w;

    int left_w = w - right_w;
    if (left_w < 0) left_w = 0;

    if (g_list)
        MoveWindow(g_list, 0, 0, left_w, content_h, TRUE);

    if (g_tracks)
        MoveWindow(g_tracks, left_w, 0, right_w, content_h, TRUE);

    if (g_btn_prev)
        MoveWindow(g_btn_prev, 6, ctrl_top + 4, 40, 24, TRUE);

    if (g_btn_play)
        MoveWindow(g_btn_play, 50, ctrl_top + 4, 52, 24, TRUE);

    if (g_btn_next)
        MoveWindow(g_btn_next, 106, ctrl_top + 4, 40, 24, TRUE);

    if (g_seekbar)
        MoveWindow(g_seekbar, 154, ctrl_top + 4, w - 290, 24, TRUE);

    if (g_status)
        MoveWindow(g_status, 0, status_top, w, status_h, TRUE);

    if (g_list) {
        MoveWindow(g_list, 0, 0, left_w, content_h, TRUE);
        ListView_Arrange(g_list, LVA_DEFAULT);
    }

    int vol_w = 100;
    if (g_vol_label)
        MoveWindow(g_vol_label, w - vol_w - 30, ctrl_top + 8, 30, 16, TRUE);

    if (g_volbar)
        MoveWindow(g_volbar,   w - vol_w - 8,  ctrl_top + 4, vol_w, 24, TRUE);

    InvalidateRect(h, &rc, TRUE);
}

static LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        InitCommonControls();
        g_hwnd = h;
        g_list = CreateWindowA(WC_LISTVIEWA, "",
            WS_CHILD | WS_VISIBLE | LVS_ICON,
            0, 0, 640, 480, h, (HMENU)1, 0, 0);
        
        g_status = CreateWindowExA(
            0, STATUSCLASSNAMEA, "",
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
            0, 0, 0, 0,
            h, (HMENU)2, GetModuleHandleA(NULL), 0);

        SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)"Ready");

        g_img = ImageList_Create(96, 96, ILC_COLOR16, 0, 32);
        ListView_SetImageList(g_list, g_img, LVSIL_NORMAL);
        SendMessageA(g_list, LVM_SETTOOLTIPS, (WPARAM)NULL, 0);
        
        g_tip = CreateWindowExA(0, TOOLTIPS_CLASSA, NULL,
            WS_POPUP | TTS_ALWAYSTIP,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            h, NULL, GetModuleHandleA(NULL), NULL);

        if (g_tip) {
            SendMessageA(g_tip, TTM_SETMAXTIPWIDTH, 0, 400);
            SendMessageA(g_tip, TTM_SETDELAYTIME, TTDT_INITIAL, 300);
            SendMessageA(g_tip, TTM_SETDELAYTIME, TTDT_AUTOPOP, 4000);

            TOOLINFOA ti;
            ZeroMemory(&ti, sizeof(ti));
            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
            ti.hwnd = h;                 // 親はメインウィンドウ
            ti.uId  = (UINT_PTR)g_list;   // 対象はListView（HWNDをIDとして使う）
            ti.lpszText = LPSTR_TEXTCALLBACKA; // 必要時に呼ぶ
            SendMessageA(g_tip, TTM_ADDTOOLA, 0, (LPARAM)&ti);
        }

        g_ui_font = CreateFontA(
            -12, 0, 0, 0, FW_NORMAL, 0, 0, 0,
            SHIFTJIS_CHARSET,
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            "MS UI Gothic"
        );
        if (g_ui_font) {
            SendMessageA(g_list,   WM_SETFONT, (WPARAM)g_ui_font, TRUE);
        }

        g_brush_btnface = CreateSolidBrush(GetSysColor(COLOR_BTNFACE));

        g_tracks = CreateWindowA(WC_LISTVIEWA, "",
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,
            0, 0, 300, 300, h, (HMENU)3, 0, 0);

        ListView_SetExtendedListViewStyle(g_tracks, LVS_EX_FULLROWSELECT);

        LVCOLUMNA col;
        ZeroMemory(&col, sizeof(col));
        col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

        col.pszText = (LPSTR)"";
        col.cx = 15;
        col.iSubItem = 0;
        ListView_InsertColumn(g_tracks, 0, &col);

        col.pszText = (LPSTR)"#";
        col.cx = 25;
        col.iSubItem = 1;
        ListView_InsertColumn(g_tracks, 1, &col);

        col.pszText = (LPSTR)"Title";
        col.cx = 170;
        col.iSubItem = 2;
        ListView_InsertColumn(g_tracks, 2, &col);

        col.pszText = (LPSTR)"Artist";
        col.cx = 115;
        col.iSubItem = 3;
        ListView_InsertColumn(g_tracks, 3, &col);

        col.pszText = (LPSTR)"Time";
        col.cx = 40;
        col.iSubItem = 4;
        ListView_InsertColumn(g_tracks, 4, &col);

        g_btn_prev = CreateWindowA(
            "BUTTON", "<<",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 40, 24,
            h, (HMENU)IDC_BTN_PREV, g_inst, NULL);

        g_btn_play = CreateWindowA(
            "BUTTON", "Play",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 52, 24,
            h, (HMENU)IDC_BTN_PLAY, g_inst, NULL);

        g_btn_next = CreateWindowA(
            "BUTTON", ">>",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 40, 24,
            h, (HMENU)IDC_BTN_NEXT, g_inst, NULL);

        g_seekbar = CreateWindowA(
            TRACKBAR_CLASSA, "",
            WS_CHILD | WS_VISIBLE | TBS_HORZ,
            0, 0, 100, 24,
            h, (HMENU)IDC_SEEKBAR, g_inst, NULL);

        g_volbar = CreateWindowA(
            TRACKBAR_CLASSA,
            "",
            WS_CHILD | WS_VISIBLE | TBS_HORZ,
            0,0,100,24,
            h,
            (HMENU)IDC_VOLBAR,
            g_inst,
            NULL);
        
        g_vol_label = CreateWindowA(
            "STATIC",
            "Vol",
            WS_CHILD | WS_VISIBLE,
            0,0,30,20,
            h,
            NULL,
            g_inst,
            NULL);

        SendMessageA(g_seekbar, TBM_SETRANGE, TRUE, MAKELONG(0, 1000));
        SendMessageA(g_seekbar, TBM_SETPOS, TRUE, 0);

        SendMessageA(g_volbar, TBM_SETPOS, TRUE, 100);

        g_old_list_proc =
            (WNDPROC)SetWindowLongA(g_list, GWL_WNDPROC, (LONG)control_subclass_proc);

        g_old_tracks_proc =
            (WNDPROC)SetWindowLongA(g_tracks, GWL_WNDPROC, (LONG)control_subclass_proc);

        g_old_seekbar_proc =
            (WNDPROC)SetWindowLongA(g_seekbar, GWL_WNDPROC, (LONG)control_subclass_proc);

        g_old_volbar_proc =
            (WNDPROC)SetWindowLongA(g_volbar, GWL_WNDPROC, (LONG)control_subclass_proc);
        
        if (g_ui_font) {
            SendMessageA(g_btn_prev, WM_SETFONT, (WPARAM)g_ui_font, TRUE);
            SendMessageA(g_btn_play, WM_SETFONT, (WPARAM)g_ui_font, TRUE);
            SendMessageA(g_btn_next, WM_SETFONT, (WPARAM)g_ui_font, TRUE);
        }

        SetTimer(h, TIMER_PLAYPOS, 300, NULL);

        PostMessageA(h, WM_APP_STARTSCAN, 0, 0);

        layout_children(h);
        return 0;
    }
    
    case WM_SIZE:
        layout_children(h);
        return 0;

    case WM_NOTIFY: {
        LPNMHDR hdr = (LPNMHDR)l;
        if (!hdr) break;

        if (hdr->hwndFrom == g_list && hdr->code == NM_DBLCLK) {
            int sel = ListView_GetNextItem(g_list, -1, LVNI_SELECTED);
            if (sel >= 0) {
                g_pending_sel = sel;
                PostMessageA(h, WM_APP_PLAYSELECT, 0, 0);
            }
            return 0;
        }

        if (hdr->hwndFrom == g_list && hdr->code == NM_CLICK) {
            int sel = ListView_GetNextItem(g_list, -1, LVNI_SELECTED);
            if (sel >= 0) {
                std::string key = album_key_from_list_index(sel);
                if (!key.empty()) {
                    g_current_album_key = key;
                    populate_tracks(key);
                }
            }
            return 0;
        }

        if (hdr->hwndFrom == g_tracks && hdr->code == NM_DBLCLK) {
            int sel = ListView_GetNextItem(g_tracks, -1, LVNI_SELECTED);
            if (sel >= 0 && !g_current_album_key.empty()) {
                play_track_in_album(g_current_album_key, sel);
            }
            return 0;
        }

        // ★ツールチップのテキスト要求
        if (hdr->hwndFrom == g_tip && hdr->code == TTN_NEEDTEXTA) {
            NMTTDISPINFOA* di = (NMTTDISPINFOA*)l;

            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(g_list, &pt);

            LVHITTESTINFO ht;
            ZeroMemory(&ht, sizeof(ht));
            ht.pt = pt;

            int idx = ListView_HitTest(g_list, &ht);
            if (idx >= 0 && (ht.flags & LVHT_ONITEM)) {

                std::string key = album_key_from_list_index(idx);
                if (!key.empty()) {
                    std::map<std::string, Album>::iterator it = g_albums.find(key);
                    if (it != g_albums.end()) {
                        Album& a = it->second;
                        int tracks = (int)a.tracks.size();

                        std::string albA = utf8_to_acp(a.album_name_u8);

                        std::string artU8 = a.albumartist_u8;

                        if (artU8.empty() && !a.tracks.empty()) {
                            if (!a.tracks[0].meta.artist.empty())
                                artU8 = a.tracks[0].meta.artist;
                        }

                        if (artU8.empty())
                            artU8 = "Unknown Artist";
                        std::string artA = utf8_to_acp(artU8);

                        if (a.year > 0) {
                            wsprintfA(g_tipbuf,
                                "%s\r\n%s\r\n%d tracks - %d",
                                albA.c_str(),
                                artA.c_str(),
                                tracks,
                                a.year);
                        } else {
                            wsprintfA(g_tipbuf,
                                "%s\r\n%s\r\n%d tracks",
                                albA.c_str(),
                                artA.c_str(),
                                tracks);
                        }

                        di->lpszText = g_tipbuf;
                        return 0;
                    }
                }
            }

            di->lpszText = (LPSTR)"";
            return 0;
        }

        break;
    }

    case WM_COMMAND:
        switch (LOWORD(w)) {
        case IDC_BTN_PREV:
            play_prev_in_album();
            return 0;

        case IDC_BTN_PLAY:
            if (can_toggle_current_stream()) {
                toggle_play_pause();
                update_playback_status();
                update_play_button_label();
            } else {
                play_default_album();
            }
            return 0;

        case IDC_BTN_NEXT:
            play_next_in_album();
            return 0;
        }
        break;

    case WM_HSCROLL:
    {
        HWND hw = (HWND)l;

        if (hw == g_volbar)
        {
            int pos = SendMessageA(g_volbar, TBM_GETPOS, 0, 0);

            BASS_SetVolume(pos);
        }

        if (hw == g_seekbar) {
            int code = LOWORD(w);

            if (code == TB_THUMBTRACK) {
                g_seek_dragging = 1;
                return 0;
            }

            if (code == TB_THUMBPOSITION || code == TB_ENDTRACK) {
                int pos = (int)SendMessageA(g_seekbar, TBM_GETPOS, 0, 0);
                seek_to_bar_pos(pos);
                g_seek_dragging = 0;
                update_playback_status();
                return 0;
            }

            if (code == TB_LINEUP || code == TB_LINEDOWN ||
                code == TB_PAGEUP || code == TB_PAGEDOWN) {
                int pos = (int)SendMessageA(g_seekbar, TBM_GETPOS, 0, 0);
                seek_to_bar_pos(pos);
                update_playback_status();
                return 0;
            }
        }
        break;
    }
    
    case WM_APP_PLAYSELECT: {
        int sel = g_pending_sel;
        g_pending_sel = -1;
        if (sel < 0) return 0;

        std::string key = album_key_from_list_index(sel);
        if (!key.empty()) {
            std::map<std::string, Album>::iterator it = g_albums.find(key);
            if (it != g_albums.end() && !it->second.tracks.empty()) {
                play_track_in_album(key, 0);
            }
        }
        return 0;
    }

    case WM_APP_PLAYENTER:
        if (!g_current_album_key.empty()) {
            play_track_in_album(g_current_album_key, (int)w);
        }
        return 0;

    case WM_APP_NEXTTRACK:
        play_next_in_album();
        return 0;

    case WM_APP_TOGGLEPLAY:
        if (can_toggle_current_stream()) {
            toggle_play_pause();
            update_playback_status();
            update_play_button_label();
        } else {
            play_default_album();
        }
        return 0;

    case WM_APP_STARTSCAN:
    {
        g_status_mode = STATUSMODE_SCANNING;
        if (!g_scan_thread) {
            DWORD tid = 0;
            g_scan_thread = CreateThread(NULL, 0, ScanThreadProc, NULL, 0, &tid);
            if (!g_scan_thread) {
                MessageBoxA(h, "CreateThread failed", "Album95", MB_OK | MB_ICONERROR);
                return 0;
            }
            SetTimer(h, TIMER_SCANPROG, 200, NULL); // ★200msごとに表示更新
        }
        return 0;
    }

    case WM_APP_SCANPROGRESS:
    {
        char buf[256];
        wsprintfA(buf, "Album95 - Scanning... %ld files", (long)w);
        SetWindowTextA(h, buf);
        return 0;
    }

    case WM_TIMER:
        if (w == TIMER_SCANPROG) {
            LONG n = g_scan_count;

            std::string f;
            EnterCriticalSection(&g_scan_cs);
            f = g_scan_file;
            LeaveCriticalSection(&g_scan_cs);

            const char* name = basenameA(f);

            char buf[512];
            wsprintfA(buf, "Scanning... %ld files   %s", (long)n, name ? name : "");
            if (g_status) SendMessageA(g_status, SB_SETTEXTA, 0, (LPARAM)buf);
            return 0;
        }
        if (w == TIMER_PLAYPOS) {
            update_playback_status();
            return 0;
        }
        break;

    case WM_APP_SCANDONE:
    {
        KillTimer(h, TIMER_SCANPROG);

        g_status_mode = STATUSMODE_BUILDING;

        regroup_by_album_with_va_fallback();
        sort_all_albums();
        rebuild_album_order();
        populate_list();

        g_status_mode = STATUSMODE_IDLE;

        char buf[256];
        wsprintfA(buf, "Album95 - %ld files", (long)g_scan_count);
        SetWindowTextA(h, buf);

        update_playback_status();
        return 0;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdc = (HDC)w;
        HWND hwCtl = (HWND)l;

        if (hwCtl == g_vol_label) {
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));
            return (LRESULT)g_brush_btnface;
        }
        break;
    }

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(h, &ps);

        RECT rc;
        GetClientRect(h, &rc);

        int w = rc.right - rc.left;
        int hgt = rc.bottom - rc.top;

        int status_h = STATUSBAR_HEIGHT;
        int ctrl_h   = CONTROLBAR_HEIGHT;

        int status_top = hgt - status_h;
        int ctrl_top   = status_top - ctrl_h;

        RECT rc_ctrl;
        rc_ctrl.left   = 0;
        rc_ctrl.top    = ctrl_top;
        rc_ctrl.right  = w;
        rc_ctrl.bottom = status_top;

        HBRUSH br = GetSysColorBrush(COLOR_BTNFACE);
        FillRect(hdc, &rc_ctrl, br);

        EndPaint(h, &ps);
        return 0;
    }

    case WM_DESTROY:
        KillTimer(h, TIMER_SCANPROG);
        KillTimer(h, TIMER_PLAYPOS);
        if (g_scan_thread) {
            WaitForSingleObject(g_scan_thread, 2000);
            CloseHandle(g_scan_thread);
            g_scan_thread = NULL;
        }
        if (g_ui_font) {
            DeleteObject(g_ui_font);
            g_ui_font = NULL;
        }
        if (g_brush_btnface) {
            DeleteObject(g_brush_btnface);
            g_brush_btnface = NULL;
        }
        BASS_Free();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcA(h, m, w, l);
}

int WINAPI WinMain(HINSTANCE h, HINSTANCE, LPSTR, int nCmdShow)
{
    g_inst = GetModuleHandleA(NULL);

    SetUnhandledExceptionFilter(CrashFilter);
    InitializeCriticalSection(&g_scan_cs);

    load_settings_from_ini();

    if (is_shift_pressed_at_startup()) {
        std::string picked;
        if (browse_for_music_folder(NULL, picked)) {
            g_root = picked;
        }

        choose_startup_sample_rate();
        save_settings_to_ini();
    }

    if (!BASS_Init(-1, (DWORD)g_bass_rate, 0, 0, NULL)) {
        MessageBoxA(NULL, "BASS_Init failed", "Album95", MB_OK);
        return 0;
    }

    BASS_SetVolume(100);

    HPLUGIN plug = BASS_PluginLoad("bass_aac.dll");
    if (!plug) {
        char buf[128];
        wsprintfA(buf, "BASS_PluginLoad(bass_aac.dll) failed. err=%d", BASS_ErrorGetCode());
        MessageBoxA(NULL, buf, "Album95", MB_OK);
    }

    WNDCLASSA wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = WndProc;
    wc.hInstance = g_inst;
    wc.lpszClassName = "Album95";
    RegisterClassA(&wc);

    HWND win = CreateWindowA(
        "Album95", "Album95",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        100, 100, 720, 520,
        0, 0, g_inst, 0);

    if (!win) {
        DWORD e = GetLastError();
        char buf[256];
        wsprintfA(buf, "CreateWindowA failed. GetLastError=%lu", (unsigned long)e);
        MessageBoxA(NULL, buf, "Album95", MB_OK | MB_ICONERROR);
        return 0;
    }

    int show = nCmdShow;
    if (GetSystemMetrics(SM_CYSCREEN) <= 600)
        show = SW_SHOWMAXIMIZED;

    ShowWindow(win, show);
    UpdateWindow(win);

    MSG msg;
    while (GetMessageA(&msg, 0, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageA(&msg);
    }
    return 0;
}
