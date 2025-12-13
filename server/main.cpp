// Minimal single-file REST-ish server with Excel COM pool.
// Compile (example, adjust Office path/SDK):
//   cl /std:c++17 main.cpp /Fe:server.exe ws2_32.lib ole32.lib oleaut32.lib
// This is a demo-quality server; hardens only lightly and assumes trusted input.

#define _WIN32_WINNT 0x0A00
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <atlbase.h>
#include <comdef.h>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <mutex>
#include <random>
#include <sstream>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#pragma comment(lib, "Ws2_32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")

namespace fs = std::filesystem;

// -------------------- Utility: base64 decode --------------------
static const std::string kBase64Alphabet =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

static const size_t kMaxHeaderLine = 8 * 1024;      // 8 KB
static const size_t kMaxHeaders = 100;               // cap header count
static const size_t kMaxBodyBytes = 5 * 1024 * 1024; // 5 MB
static const int kListenBacklog = SOMAXCONN;

// Forward declarations for helpers
std::string trim(const std::string &s);
std::vector<std::string> split_csv(const std::string &csv);

std::string base64_decode(const std::string &input) {
    std::string out;
    std::vector<int> T(256, -1);
    for (size_t i = 0; i < kBase64Alphabet.size(); i++) {
        T[static_cast<unsigned char>(kBase64Alphabet[i])] = static_cast<int>(i);
    }
    int val = 0, valb = -8;
    for (unsigned char c : input) {
        if (T[c] == -1) break;
        val = (val << 6) + T[c];
        valb += 6;
        if (valb >= 0) {
            out.push_back(char((val >> valb) & 0xFF));
            valb -= 8;
        }
    }
    return out;
}

std::string base64_encode(const std::string &input) {
    std::string out;
    int val = 0;
    int valb = -6;
    for (unsigned char c : input) {
        val = (val << 8) + c;
        valb += 8;
        while (valb >= 0) {
            out.push_back(kBase64Alphabet[(val >> valb) & 0x3F]);
            valb -= 6;
        }
    }
    if (valb > -6) out.push_back(kBase64Alphabet[((val << 8) >> (valb + 8)) & 0x3F]);
    while (out.size() % 4) out.push_back('=');
    return out;
}

// -------------------- Utility: naive JSON extraction --------------------
// Very small helper to pull primitive values from a flat JSON object.
std::string extract_json_string(const std::string &body, const std::string &key) {
    std::string token = "\"" + key + "\"";
    size_t key_pos = body.find(token);
    if (key_pos == std::string::npos) return "";
    size_t colon = body.find(':', key_pos);
    if (colon == std::string::npos) return "";
    size_t start = body.find('"', colon);
    if (start == std::string::npos) return "";
    ++start; // move past opening quote
    std::string result;
    bool escape = false;
    for (size_t i = start; i < body.size(); ++i) {
        char ch = body[i];
        if (escape) {
            switch (ch) {
                case '"': result.push_back('"'); break;
                case '\\': result.push_back('\\'); break;
                case '/': result.push_back('/'); break;
                case 'b': result.push_back('\b'); break;
                case 'f': result.push_back('\f'); break;
                case 'n': result.push_back('\n'); break;
                case 'r': result.push_back('\r'); break;
                case 't': result.push_back('\t'); break;
                default: result.push_back(ch); break;
            }
            escape = false;
            continue;
        }
        if (ch == '\\') {
            escape = true;
            continue;
        }
        if (ch == '"') {
            return result;
        }
        result.push_back(ch);
    }
    return "";
}

int extract_json_int(const std::string &body, const std::string &key, int def = 0) {
    std::string token = "\"" + key + "\"";
    size_t pos = body.find(token);
    if (pos == std::string::npos) return def;
    pos = body.find(':', pos);
    if (pos == std::string::npos) return def;
    size_t start = body.find_first_of("-0123456789", pos);
    if (start == std::string::npos) return def;
    size_t end = start;
    while (end < body.size() && isdigit(body[end])) end++;
    return std::stoi(body.substr(start, end - start));
}

bool extract_json_bool(const std::string &body, const std::string &key, bool def = false) {
    std::string token = "\"" + key + "\"";
    size_t pos = body.find(token);
    if (pos == std::string::npos) return def;
    pos = body.find(':', pos);
    if (pos == std::string::npos) return def;
    size_t start = body.find_first_not_of(" \t\r\n", pos + 1);
    if (start == std::string::npos) return def;
    if (body.compare(start, 4, "true") == 0) return true;
    if (body.compare(start, 5, "false") == 0) return false;
    return def;
}

double extract_json_double(const std::string &body, const std::string &key, double def = 0.0) {
    std::string token = "\"" + key + "\"";
    size_t pos = body.find(token);
    if (pos == std::string::npos) return def;
    pos = body.find(':', pos);
    if (pos == std::string::npos) return def;
    size_t start = body.find_first_of("-0123456789", pos);
    if (start == std::string::npos) return def;
    size_t end = start;
    while (end < body.size() && (isdigit(body[end]) || body[end] == '.')) end++;
    try {
        return std::stod(body.substr(start, end - start));
    } catch (...) {
        return def;
    }
}

bool json_has_key(const std::string &body, const std::string &key) {
    return body.find("\"" + key + "\"") != std::string::npos;
}

// Forward declarations
std::string variant_to_json(const VARIANT &v);

// -------------------- Config --------------------
struct Config {
    int port = 8080;
    int excel_instances = 1;
    std::unordered_map<std::string, std::string> users; // username -> password
    std::unordered_set<std::string> admins;             // admin usernames from config only
};

Config load_config(const std::string &path) {
    Config cfg;
    std::ifstream in(path);
    if (!in.is_open()) {
        cfg.users["admin"] = "admin";
        cfg.admins.insert("admin");
        return cfg;
    }
    std::stringstream buffer;
    buffer << in.rdbuf();
    std::string body = buffer.str();
    cfg.port = extract_json_int(body, "port", 8080);
    cfg.excel_instances = extract_json_int(body, "excel_instances", 1);
    // Users: expects [{"username":"u","password":"p"}]
    size_t pos = 0;
    while ((pos = body.find("\"username\"", pos)) != std::string::npos) {
        std::string user = extract_json_string(body.substr(pos), "username");
        std::string pass = extract_json_string(body.substr(pos), "password");
        if (!user.empty()) cfg.users[user] = pass;
        pos += 10;
    }
    // Admins: expects ["name1","name2"] under key "admins"
    size_t a = body.find("\"admins\"");
    if (a != std::string::npos) {
        size_t lb = body.find('[', a);
        size_t rb = body.find(']', lb);
        if (lb != std::string::npos && rb != std::string::npos && rb > lb) {
            size_t cursor = lb;
            while (true) {
                size_t q1 = body.find('"', cursor);
                if (q1 == std::string::npos || q1 >= rb) break;
                size_t q2 = body.find('"', q1 + 1);
                if (q2 == std::string::npos || q2 > rb) break;
                std::string name = body.substr(q1 + 1, q2 - q1 - 1);
                if (!name.empty()) cfg.admins.insert(name);
                cursor = q2 + 1;
            }
        }
    }
    if (cfg.users.empty()) cfg.users["admin"] = "admin";
    if (cfg.admins.empty()) cfg.admins.insert("admin");
    return cfg;
}

// -------------------- Utility: string helpers --------------------
std::string trim(const std::string &s) {
    size_t b = s.find_first_not_of(" \t\r\n");
    if (b == std::string::npos) return "";
    size_t e = s.find_last_not_of(" \t\r\n");
    return s.substr(b, e - b + 1);
}

std::vector<std::string> split_csv(const std::string &csv) {
    std::vector<std::string> out;
    std::stringstream ss(csv);
    std::string item;
    while (std::getline(ss, item, ',')) {
        item = trim(item);
        if (!item.empty()) out.push_back(item);
    }
    return out;
}

// -------------------- Binary database --------------------
enum class Role : uint32_t { User = 0, Developer = 1, Admin = 2 };

struct UserRecord {
    std::string name;
    std::string password;
    std::vector<std::string> groups;
    Role role = Role::User;
};

struct AppRecord {
    std::string owner;
    std::string name;
    int latest_version = 1;
    std::string description;
    bool public_access = true;
    std::string access_group; // if not public, gate by this group
};

std::string app_key(const std::string &owner, const std::string &name) {
    return owner + "/" + name;
}

class Database {
  public:
    explicit Database(std::string path) : path_(std::move(path)) {}

    bool load() {
        std::lock_guard<std::mutex> lock(mu_);
        std::ifstream in(path_, std::ios::binary);
        if (!in.is_open()) {
            // Create default admin user
            UserRecord admin{"admin", "admin", {"admin"}};
            users_[admin.name] = admin;
            return save_locked();
        }
        char magic[4];
        in.read(magic, 4);
        if (std::string(magic, 4) != "DB01") return false;
        uint32_t version = 0;
        in.read(reinterpret_cast<char *>(&version), sizeof(version));
        if (version != 1 && version != 2) return false;
        auto read_string = [&in](std::string &out) {
            uint32_t len = 0;
            in.read(reinterpret_cast<char *>(&len), sizeof(len));
            out.resize(len);
            if (len > 0) in.read(&out[0], len);
        };
        uint32_t user_count = 0;
        in.read(reinterpret_cast<char *>(&user_count), sizeof(user_count));
        for (uint32_t i = 0; i < user_count; ++i) {
            UserRecord u;
            read_string(u.name);
            read_string(u.password);
            uint32_t gcount = 0;
            in.read(reinterpret_cast<char *>(&gcount), sizeof(gcount));
            for (uint32_t g = 0; g < gcount; ++g) {
                std::string gname;
                read_string(gname);
                u.groups.push_back(gname);
            }
            if (version == 2) {
                uint32_t role_val = 0;
                in.read(reinterpret_cast<char *>(&role_val), sizeof(role_val));
                if (role_val <= 2) u.role = static_cast<Role>(role_val);
            }
            users_[u.name] = u;
        }
        uint32_t app_count = 0;
        in.read(reinterpret_cast<char *>(&app_count), sizeof(app_count));
        for (uint32_t i = 0; i < app_count; ++i) {
            AppRecord a;
            read_string(a.owner);
            read_string(a.name);
            in.read(reinterpret_cast<char *>(&a.latest_version), sizeof(a.latest_version));
            read_string(a.description);
            uint8_t pub = 1;
            in.read(reinterpret_cast<char *>(&pub), sizeof(pub));
            a.public_access = pub != 0;
            read_string(a.access_group);
            apps_[app_key(a.owner, a.name)] = a;
        }
        return true;
    }

    bool save() {
        std::lock_guard<std::mutex> lock(mu_);
        return save_locked();
    }

    bool get_user(const std::string &name, UserRecord &out) {
        std::lock_guard<std::mutex> lock(mu_);
        auto it = users_.find(name);
        if (it == users_.end()) return false;
        out = it->second;
        return true;
    }

    bool upsert_user(const UserRecord &u) {
        std::lock_guard<std::mutex> lock(mu_);
        users_[u.name] = u;
        return save_locked();
    }

    bool list_users(std::vector<UserRecord> &out) {
        std::lock_guard<std::mutex> lock(mu_);
        out.clear();
        for (auto &kv : users_) out.push_back(kv.second);
        return true;
    }

    bool upsert_app(const AppRecord &a) {
        std::lock_guard<std::mutex> lock(mu_);
        apps_[app_key(a.owner, a.name)] = a;
        return save_locked();
    }

    bool get_app(const std::string &owner, const std::string &name, AppRecord &out) {
        std::lock_guard<std::mutex> lock(mu_);
        auto it = apps_.find(app_key(owner, name));
        if (it == apps_.end()) return false;
        out = it->second;
        return true;
    }

    bool remove_app(const std::string &owner, const std::string &name) {
        std::lock_guard<std::mutex> lock(mu_);
        apps_.erase(app_key(owner, name));
        return save_locked();
    }

    std::vector<AppRecord> list_apps() {
        std::lock_guard<std::mutex> lock(mu_);
        std::vector<AppRecord> out;
        out.reserve(apps_.size());
        for (auto &kv : apps_) out.push_back(kv.second);
        return out;
    }

  private:
    bool save_locked() {
        std::string tmp = path_ + ".tmp";
        std::ofstream out(tmp, std::ios::binary | std::ios::trunc);
        if (!out.is_open()) return false;
        auto write_string = [&out](const std::string &s) {
            uint32_t len = static_cast<uint32_t>(s.size());
            out.write(reinterpret_cast<const char *>(&len), sizeof(len));
            if (len) out.write(s.data(), len);
        };
        out.write("DB01", 4);
        uint32_t version = 2;
        out.write(reinterpret_cast<const char *>(&version), sizeof(version));
        uint32_t user_count = static_cast<uint32_t>(users_.size());
        out.write(reinterpret_cast<const char *>(&user_count), sizeof(user_count));
        for (auto &kv : users_) {
            const UserRecord &u = kv.second;
            write_string(u.name);
            write_string(u.password);
            uint32_t gcount = static_cast<uint32_t>(u.groups.size());
            out.write(reinterpret_cast<const char *>(&gcount), sizeof(gcount));
            for (auto &g : u.groups) write_string(g);
            uint32_t role_val = static_cast<uint32_t>(u.role);
            out.write(reinterpret_cast<const char *>(&role_val), sizeof(role_val));
        }
        uint32_t app_count = static_cast<uint32_t>(apps_.size());
        out.write(reinterpret_cast<const char *>(&app_count), sizeof(app_count));
        for (auto &kv : apps_) {
            const AppRecord &a = kv.second;
            write_string(a.owner);
            write_string(a.name);
            out.write(reinterpret_cast<const char *>(&a.latest_version), sizeof(a.latest_version));
            write_string(a.description);
            uint8_t pub = a.public_access ? 1 : 0;
            out.write(reinterpret_cast<const char *>(&pub), sizeof(pub));
            write_string(a.access_group);
        }
        out.close();
        std::error_code ec;
        fs::rename(tmp, path_, ec);
        if (ec) return false;
        return true;
    }

    std::string path_;
    std::unordered_map<std::string, UserRecord> users_;
    std::unordered_map<std::string, AppRecord> apps_;
    std::mutex mu_;
};

// -------------------- Excel Pool --------------------
bool dispatch_put_bool(IDispatch *disp, const wchar_t *name, bool value) {
    if (!disp) return false;
    DISPID dispid;
    LPOLESTR names[1];
    names[0] = const_cast<LPOLESTR>(name);
    if (FAILED(disp->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &dispid))) return false;
    VARIANTARG arg;
    VariantInit(&arg);
    arg.vt = VT_BOOL;
    arg.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    DISPPARAMS params{};
    params.rgvarg = &arg;
    params.cArgs = 1;
    DISPID named = DISPID_PROPERTYPUT;
    params.rgdispidNamedArgs = &named;
    params.cNamedArgs = 1;
    return SUCCEEDED(disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, nullptr, nullptr, nullptr));
}

void dispatch_call_noargs(IDispatch *disp, const wchar_t *name) {
    if (!disp) return;
    DISPID dispid;
    LPOLESTR names[1];
    names[0] = const_cast<LPOLESTR>(name);
    if (FAILED(disp->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &dispid))) return;
    DISPPARAMS params{};
    disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, nullptr, nullptr, nullptr);
}

bool dispatch_invoke(IDispatch *disp, const wchar_t *name, WORD flags, VARIANT *args, UINT cargs, VARIANT *result) {
    if (!disp) return false;
    DISPID dispid;
    LPOLESTR names[1];
    names[0] = const_cast<LPOLESTR>(name);
    if (FAILED(disp->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &dispid))) return false;
    DISPPARAMS params{};
    params.rgvarg = args;
    params.cArgs = cargs;
    DISPID named = DISPID_PROPERTYPUT;
    if (flags & DISPATCH_PROPERTYPUT) {
        params.rgdispidNamedArgs = &named;
        params.cNamedArgs = 1;
    }
    return SUCCEEDED(disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags, &params, result, nullptr, nullptr));
}

CComPtr<IDispatch> dispatch_get(IDispatch *disp, const wchar_t *name) {
    VARIANT res;
    VariantInit(&res);
    if (!dispatch_invoke(disp, name, DISPATCH_PROPERTYGET, nullptr, 0, &res)) return nullptr;
    if (res.vt == VT_DISPATCH) {
        CComPtr<IDispatch> out = res.pdispVal;
        return out;
    }
    VariantClear(&res);
    return nullptr;
}

CComPtr<IDispatch> dispatch_call_bstr(IDispatch *disp, const wchar_t *name, const std::wstring &arg) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocString(arg.c_str());
    VARIANT res;
    VariantInit(&res);
    bool ok = dispatch_invoke(disp, name, DISPATCH_METHOD, &v, 1, &res);
    VariantClear(&v);
    if (!ok) return nullptr;
    if (res.vt == VT_DISPATCH) {
        CComPtr<IDispatch> out = res.pdispVal;
        return out;
    }
    VariantClear(&res);
    return nullptr;
}

bool dispatch_put_variant(IDispatch *disp, const wchar_t *name, VARIANT *val) {
    return dispatch_invoke(disp, name, DISPATCH_PROPERTYPUT, val, 1, nullptr);
}

class ExcelPool {
  public:
    bool init(int count) {
        HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
            std::cerr << "COM init failed\n";
            return false;
        }
        com_initialized_ = (hr == S_OK || hr == S_FALSE);
        for (int i = 0; i < count; ++i) {
            CComPtr<IDispatch> app = create_instance();
            if (!app) return false;
            Slot slot;
            slot.app = app;
            slots_.push_back(slot);
        }
        return true;
    }

    void shutdown() {
        for (auto &slot : slots_) {
            try {
                if (slot.workbook) dispatch_call_noargs(slot.workbook, L"Close");
                dispatch_call_noargs(slot.app, L"Quit");
            } catch (...) {
            }
        }
        slots_.clear();
        if (com_initialized_) CoUninitialize();
    }

    bool load_workbook(const std::string &session_id, const std::string &user, const fs::path &path, std::string &err) {
        int slot_index = -1;
        CComPtr<IDispatch> app;
        CComPtr<IDispatch> old_wb;
        {
            std::lock_guard<std::mutex> lock(mu_);
            slot_index = find_or_acquire_slot_locked(session_id, user);
            if (slot_index < 0) {
                err = "no available excel instances";
                return false;
            }
            app = slots_[slot_index].app;
            old_wb = slots_[slot_index].workbook;
        }
        if (old_wb) {
            dispatch_call_noargs(old_wb, L"Close");
        }
        std::wstring wpath = path.wstring();
        CComPtr<IDispatch> workbooks = dispatch_get(app, L"Workbooks");
        if (!workbooks) {
            err = "failed to reach Workbooks";
            std::lock_guard<std::mutex> lock(mu_);
            release_slot_locked(session_id);
            return false;
        }
        CComPtr<IDispatch> workbook = dispatch_call_bstr(workbooks, L"Open", wpath);
        if (!workbook) {
            err = "failed to open workbook";
            std::lock_guard<std::mutex> lock(mu_);
            release_slot_locked(session_id);
            return false;
        }
        dispatch_put_bool(app, L"DisplayAlerts", false);
        {
            std::lock_guard<std::mutex> lock(mu_);
            slots_[slot_index].workbook = workbook;
            slots_[slot_index].workbook_path = path;
            slots_[slot_index].session_id = session_id;
            slots_[slot_index].user = user;
            slots_[slot_index].in_use = true;
        }
        return true;
    }

    bool query_range(const std::string &session_id, const std::string &sheet, const std::string &range, std::string &json_out, std::string &err) {
        CComPtr<IDispatch> range_obj;
        if (!get_range(session_id, sheet, range, range_obj, err)) return false;
        VARIANT res;
        VariantInit(&res);
        if (!dispatch_invoke(range_obj, L"Value", DISPATCH_PROPERTYGET, nullptr, 0, &res)) {
            err = "failed to read value";
            return false;
        }
        json_out = variant_to_json(res);
        VariantClear(&res);
        return true;
    }

    bool set_range_value(const std::string &session_id, const std::string &sheet, const std::string &range, VARIANT &val, std::string &err) {
        CComPtr<IDispatch> range_obj;
        if (!get_range(session_id, sheet, range, range_obj, err)) return false;
        if (!dispatch_put_variant(range_obj, L"Value", &val)) {
            err = "failed to set value";
            return false;
        }
        return true;
    }

    bool close_session(const std::string &session_id, bool restart, std::string &err) {
        int idx = -1;
        CComPtr<IDispatch> wb;
        {
            std::lock_guard<std::mutex> lock(mu_);
            idx = find_slot_locked(session_id);
            if (idx < 0) {
                err = "no active session";
                return false;
            }
            wb = slots_[idx].workbook;
        }
        if (wb) dispatch_call_noargs(wb, L"Close");
        {
            std::lock_guard<std::mutex> lock(mu_);
            slots_[idx].workbook.Release();
            slots_[idx].workbook_path.clear();
            slots_[idx].session_id.clear();
            slots_[idx].user.clear();
            slots_[idx].in_use = false;
            if (restart && !restart_slot_locked(static_cast<size_t>(idx))) {
                err = "failed to restart excel";
                return false;
            }
        }
        return true;
    }

  private:
    struct Slot {
        CComPtr<IDispatch> app;
        CComPtr<IDispatch> workbook;
        std::string session_id;
        std::string user;
        fs::path workbook_path;
        bool in_use = false;
    };

    CComPtr<IDispatch> create_instance() {
        CLSID clsid;
        if (FAILED(::CLSIDFromProgID(L"Excel.Application", &clsid))) {
            std::cerr << "Excel not registered\n";
            return nullptr;
        }
        CComPtr<IDispatch> app;
        if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, reinterpret_cast<void **>(&app)))) {
            std::cerr << "Excel instance creation failed\n";
            return nullptr;
        }
        dispatch_put_bool(app, L"Visible", false);
        dispatch_put_bool(app, L"DisplayAlerts", false);
        return app;
    }

    int find_slot_locked(const std::string &session_id) {
        for (size_t i = 0; i < slots_.size(); ++i) {
            if (slots_[i].session_id == session_id) return static_cast<int>(i);
        }
        return -1;
    }

    int find_or_acquire_slot_locked(const std::string &session_id, const std::string &user) {
        int existing = find_slot_locked(session_id);
        if (existing >= 0) return existing;
        for (size_t i = 0; i < slots_.size(); ++i) {
            if (!slots_[i].in_use) {
                slots_[i].in_use = true;
                slots_[i].session_id = session_id;
                slots_[i].user = user;
                slots_[i].workbook.Release();
                slots_[i].workbook_path.clear();
                return static_cast<int>(i);
            }
        }
        return -1;
    }

    bool restart_slot_locked(size_t idx) {
        if (idx >= slots_.size()) return false;
        if (slots_[idx].app) dispatch_call_noargs(slots_[idx].app, L"Quit");
        slots_[idx].app.Release();
        slots_[idx].workbook.Release();
        CComPtr<IDispatch> app = create_instance();
        if (!app) return false;
        slots_[idx].app = app;
        return true;
    }

    bool get_range(const std::string &session_id, const std::string &sheet, const std::string &range, CComPtr<IDispatch> &range_out, std::string &err) {
        CComPtr<IDispatch> wb;
        {
            std::lock_guard<std::mutex> lock(mu_);
            int idx = find_slot_locked(session_id);
            if (idx < 0 || !slots_[idx].workbook) {
                err = "no workbook loaded";
                return false;
            }
            wb = slots_[idx].workbook;
        }
        CComPtr<IDispatch> sheets = dispatch_get(wb, L"Worksheets");
        if (!sheets) {
            err = "worksheets not available";
            return false;
        }
        CComPtr<IDispatch> sheet_obj = dispatch_call_bstr(sheets, L"Item", std::wstring(sheet.begin(), sheet.end()));
        if (!sheet_obj) {
            err = "sheet not found";
            return false;
        }
        CComPtr<IDispatch> rng = dispatch_call_bstr(sheet_obj, L"Range", std::wstring(range.begin(), range.end()));
        if (!rng) {
            err = "range not found";
            return false;
        }
        range_out = rng;
        return true;
    }

    void release_slot_locked(const std::string &session_id) {
        int idx = find_slot_locked(session_id);
        if (idx < 0) return;
        slots_[idx].workbook.Release();
        slots_[idx].workbook_path.clear();
        slots_[idx].session_id.clear();
        slots_[idx].user.clear();
        slots_[idx].in_use = false;
    }

    std::vector<Slot> slots_;
    bool com_initialized_ = false;
    std::mutex mu_;
};

// -------------------- Session store --------------------
class SessionStore {
  public:
    std::string login(const std::string &user) {
        std::lock_guard<std::mutex> lock(mu_);
        std::string token = generate_token();
        sessions_[token] = user;
        return token;
    }

    void logout(const std::string &token) {
        std::lock_guard<std::mutex> lock(mu_);
        sessions_.erase(token);
    }

    bool verify(const std::string &token, std::string &user_out) {
        std::lock_guard<std::mutex> lock(mu_);
        auto it = sessions_.find(token);
        if (it == sessions_.end()) return false;
        user_out = it->second;
        return true;
    }

  private:
    std::string generate_token() {
        static const char charset[] = "abcdefghijklmnopqrstuvwxyz0123456789";
        thread_local std::mt19937 rng{std::random_device{}()};
        std::uniform_int_distribution<int> dist(0, static_cast<int>(sizeof(charset) - 2));
        std::string t(40, ' ');
        for (char &c : t) c = charset[dist(rng)];
        return t;
    }
    std::unordered_map<std::string, std::string> sessions_;
    std::mutex mu_;
};

// -------------------- HTTP primitives --------------------
struct HttpRequest {
    std::string method;
    std::string path;
    std::unordered_map<std::string, std::string> headers;
    std::string body;
};

struct HttpResponse {
    int status = 200;
    std::string content_type = "application/json";
    std::string body = "{}";
};

std::string read_line(SOCKET s) {
    std::string line;
    char c;
    while (recv(s, &c, 1, 0) == 1) {
        if (c == '\r') {
            recv(s, &c, 1, 0); // consume \n
            break;
        }
        line.push_back(c);
        if (line.size() > kMaxHeaderLine) return ""; // line too long
    }
    return line;
}

bool read_request(SOCKET s, HttpRequest &req) {
    std::string request_line = read_line(s);
    if (request_line.empty()) return false;
    std::istringstream rl(request_line);
    rl >> req.method >> req.path;
    if (req.method.empty() || req.path.empty()) return false;
    size_t header_count = 0;
    while (true) {
        std::string header_line = read_line(s);
        if (header_line.empty()) break;
        if (++header_count > kMaxHeaders) return false;
        size_t colon = header_line.find(':');
        if (colon == std::string::npos) continue;
        std::string key = header_line.substr(0, colon);
        std::string val = header_line.substr(colon + 1);
        while (!val.empty() && (val.front() == ' ' || val.front() == '\t')) val.erase(val.begin());
        req.headers[key] = val;
    }
    auto it = req.headers.find("Content-Length");
    if (it != req.headers.end()) {
        long long len = 0;
        try {
            len = std::stoll(it->second);
        } catch (...) {
            return false;
        }
        if (len < 0 || static_cast<size_t>(len) > kMaxBodyBytes) return false;
        req.body.resize(static_cast<size_t>(len));
        size_t received = 0;
        while (received < static_cast<size_t>(len)) {
            int r = recv(s, &req.body[received], static_cast<int>(len - received), 0);
            if (r <= 0) return false;
            received += static_cast<size_t>(r);
        }
    }
    return true;
}

void send_response(SOCKET s, const HttpResponse &resp) {
    std::ostringstream oss;
    oss << "HTTP/1.1 " << resp.status << "\r\n";
    oss << "Content-Type: " << resp.content_type << "\r\n";
    oss << "Content-Length: " << resp.body.size() << "\r\n";
    oss << "Access-Control-Allow-Origin: *\r\n";
    oss << "Access-Control-Allow-Headers: Content-Type, Authorization\r\n";
    oss << "Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS\r\n";
    oss << "Connection: close\r\n\r\n";
    oss << resp.body;
    std::string data = oss.str();
    send(s, data.c_str(), static_cast<int>(data.size()), 0);
}

// -------------------- App management --------------------
struct AppInfo {
    std::string name;
    int version = 1;
    std::string description;
};

fs::path app_root() {
    return fs::path("app");
}

bool ensure_dir(const fs::path &p) {
    std::error_code ec;
    if (fs::exists(p, ec)) return true;
    return fs::create_directories(p, ec);
}

bool is_safe_name(const std::string &name) {
    if (name.empty() || name.size() > 128) return false;
    for (char c : name) {
        if (!(isalnum(static_cast<unsigned char>(c)) || c == '_' || c == '-' || c == '.')) return false;
    }
    if (name.find("..") != std::string::npos) return false;
    return name.find_first_of("/\\") == std::string::npos;
}

fs::path version_path(const std::string &owner, const std::string &app, int version) {
    return app_root() / owner / app / std::to_string(version);
}

fs::path app_image_path(const std::string &owner, const std::string &app) {
    return app_root() / owner / app / "cover.png";
}

fs::path app_image_version_path(const std::string &owner, const std::string &app, int version) {
    return version_path(owner, app, version) / "cover.png";
}

bool save_app_image(const std::string &owner, const std::string &app, const std::string &image_b64) {
    if (image_b64.empty()) return true;
    std::string bytes = base64_decode(image_b64);
    if (bytes.empty()) return false;
    fs::path img_path = app_image_path(owner, app);
    if (!ensure_dir(img_path.parent_path())) return false;
    std::ofstream out(img_path, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) return false;
    out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    return true;
}

bool save_app_image_version(const std::string &owner, const std::string &app, int version, const std::string &image_b64) {
    if (image_b64.empty()) return true;
    std::string bytes = base64_decode(image_b64);
    if (bytes.empty()) return false;
    fs::path img_path = app_image_version_path(owner, app, version);
    if (!ensure_dir(img_path.parent_path())) return false;
    std::ofstream out(img_path, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) return false;
    out.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    return true;
}

bool copy_app_image_version(const std::string &owner, const std::string &app, int from_version, int to_version) {
    fs::path src = app_image_version_path(owner, app, from_version);
    if (!fs::exists(src)) src = app_image_path(owner, app);
    if (!fs::exists(src)) return false;
    fs::path dst = app_image_version_path(owner, app, to_version);
    if (!ensure_dir(dst.parent_path())) return false;
    std::error_code ec;
    fs::copy_file(src, dst, fs::copy_options::overwrite_existing, ec);
    return !ec;
}

std::string load_app_image_base64(const std::string &owner, const std::string &app) {
    fs::path img_path = app_image_path(owner, app);
    std::ifstream in(img_path, std::ios::binary);
    if (!in.is_open()) return "";
    std::ostringstream buffer;
    buffer << in.rdbuf();
    return base64_encode(buffer.str());
}

std::string load_app_image_base64_version(const std::string &owner, const std::string &app, int version) {
    fs::path img_path = app_image_version_path(owner, app, version);
    std::ifstream in(img_path, std::ios::binary);
    if (!in.is_open()) return "";
    std::ostringstream buffer;
    buffer << in.rdbuf();
    return base64_encode(buffer.str());
}

fs::path app_ui_path(const std::string &owner, const std::string &app) {
    return app_root() / owner / app / "ui.json";
}

fs::path app_ui_version_path(const std::string &owner, const std::string &app, int version) {
    return version_path(owner, app, version) / "ui.json";
}

bool save_app_ui(const std::string &owner, const std::string &app, const std::string &json) {
    fs::path path = app_ui_path(owner, app);
    if (!ensure_dir(path.parent_path())) return false;
    std::ofstream out(path, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) return false;
    out << json;
    return true;
}

bool save_app_ui_version(const std::string &owner, const std::string &app, int version, const std::string &json) {
    fs::path path = app_ui_version_path(owner, app, version);
    if (!ensure_dir(path.parent_path())) return false;
    std::ofstream out(path, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) return false;
    out << json;
    return true;
}

std::string load_app_ui(const std::string &owner, const std::string &app) {
    fs::path path = app_ui_path(owner, app);
    std::ifstream in(path, std::ios::binary);
    if (!in.is_open()) return "";
    std::ostringstream buffer;
    buffer << in.rdbuf();
    return buffer.str();
}

std::string load_app_ui_version(const std::string &owner, const std::string &app, int version) {
    fs::path path = app_ui_version_path(owner, app, version);
    std::ifstream in(path, std::ios::binary);
    if (!in.is_open()) return "";
    std::ostringstream buffer;
    buffer << in.rdbuf();
    return buffer.str();
}

bool write_metadata(const fs::path &p, const AppInfo &info) {
    std::ofstream out(p / "meta.txt", std::ios::trunc);
    if (!out) return false;
    out << "name=" << info.name << "\n";
    out << "version=" << info.version << "\n";
    out << "description=" << info.description << "\n";
    return true;
}

std::string json_escape(const std::string &s) {
    std::string out;
    for (char c : s) {
        if (c == '"' || c == '\\') out.push_back('\\');
        out.push_back(c);
    }
    return out;
}

std::string variant_to_json(const VARIANT &v) {
    VARIANT val;
    VariantInit(&val);
    VariantCopy(&val, const_cast<VARIANT *>(&v));
    if (val.vt == VT_EMPTY || val.vt == VT_NULL) return "null";
    if (val.vt == VT_BOOL) return val.boolVal == VARIANT_TRUE ? "true" : "false";
    if (val.vt == VT_I4 || val.vt == VT_INT) return std::to_string(val.intVal);
    if (val.vt == VT_R8) return std::to_string(val.dblVal);
    if (val.vt == VT_BSTR) {
        std::wstring ws(val.bstrVal ? val.bstrVal : L"");
        std::string s(ws.begin(), ws.end());
        VariantClear(&val);
        return "\"" + json_escape(s) + "\"";
    }
    if ((val.vt & VT_ARRAY) && (val.vt & VT_VARIANT) && val.parray) {
        SAFEARRAY *arr = val.parray;
        LONG lbound1 = 0, ubound1 = -1, lbound2 = 0, ubound2 = -1;
        UINT dims = SafeArrayGetDim(arr);
        SafeArrayGetLBound(arr, 1, &lbound1);
        SafeArrayGetUBound(arr, 1, &ubound1);
        if (dims >= 2) {
            SafeArrayGetLBound(arr, 2, &lbound2);
            SafeArrayGetUBound(arr, 2, &ubound2);
        }
        std::ostringstream oss;
        if (dims == 1) {
            oss << "[";
            bool first = true;
            for (LONG i = lbound1; i <= ubound1; ++i) {
                VARIANT item;
                VariantInit(&item);
                LONG idx = i;
                if (SUCCEEDED(SafeArrayGetElement(arr, &idx, &item))) {
                    if (!first) oss << ",";
                    first = false;
                    oss << variant_to_json(item);
                }
                VariantClear(&item);
            }
            oss << "]";
        } else if (dims == 2) {
            oss << "[";
            bool first_row = true;
            for (LONG r = lbound1; r <= ubound1; ++r) {
                if (!first_row) oss << ",";
                first_row = false;
                oss << "[";
                bool first_col = true;
                for (LONG c = lbound2; c <= ubound2; ++c) {
                    VARIANT item;
                    VariantInit(&item);
                    LONG idx[2] = {r, c};
                    if (SUCCEEDED(SafeArrayGetElement(arr, idx, &item))) {
                        if (!first_col) oss << ",";
                        first_col = false;
                        oss << variant_to_json(item);
                    }
                    VariantClear(&item);
                }
                oss << "]";
            }
            oss << "]";
        } else {
            VariantClear(&val);
            return "null";
        }
        VariantClear(&val);
        return oss.str();
    }
    VariantClear(&val);
    return "null";
}

bool user_in_group(const UserRecord &u, const std::string &group) {
    for (auto &g : u.groups) {
        if (g == group) return true;
    }
    return false;
}

bool is_admin(const UserRecord &u, const Config &cfg) {
    return cfg.admins.count(u.name) > 0;
}

bool can_access(const AppRecord &app, const UserRecord &u) {
    if (app.public_access) return true;
    if (app.owner == u.name) return true;
    if (!app.access_group.empty() && user_in_group(u, app.access_group)) return true;
    return false;
}

// -------------------- Handlers --------------------
class Server {
    public:
        Server(const Config &cfg, ExcelPool &pool, Database &db) : cfg_(cfg), pool_(pool), db_(db) {}

    void start() {
        WSADATA wsa;
        if (WSAStartup(MAKEWORD(2, 2), &wsa) != 0) {
            std::cerr << "WSAStartup failed\n";
            return;
        }
        SOCKET listen_socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
        if (listen_socket == INVALID_SOCKET) {
            std::cerr << "Socket creation failed\n";
            return;
        }
        sockaddr_in service{};
        service.sin_family = AF_INET;
        service.sin_addr.s_addr = INADDR_ANY;
        service.sin_port = htons(static_cast<u_short>(cfg_.port));
        if (bind(listen_socket, reinterpret_cast<SOCKADDR *>(&service), sizeof(service)) == SOCKET_ERROR) {
            std::cerr << "Bind failed\n";
            closesocket(listen_socket);
            return;
        }
        listen(listen_socket, kListenBacklog);
        std::cout << "Server listening on port " << cfg_.port << "\n";
        running_ = true;
        while (running_) {
            SOCKET client = accept(listen_socket, nullptr, nullptr);
            if (client == INVALID_SOCKET) continue;
            std::thread(&Server::handle_client, this, client).detach();
        }
        closesocket(listen_socket);
        WSACleanup();
    }

  private:
    void handle_client(SOCKET client) {
        HttpRequest req;
        HttpResponse resp;
        if (!read_request(client, req)) {
            resp.status = 400;
            resp.body = "{\"error\":\"bad request\"}";
        } else {
            resp = dispatch(req);
        }
        send_response(client, resp);
        closesocket(client);
    }

    HttpResponse dispatch(const HttpRequest &req) {
        if (req.method == "OPTIONS") {
            HttpResponse resp;
            resp.status = 200;
            resp.body = "";
            return resp;
        }
        if (req.method == "GET" && req.path == "/health") return handle_health(req);
        if (req.method == "POST" && req.path == "/login") return handle_login(req);
        if (req.method == "POST" && req.path == "/logout") return handle_logout(req);
        if (req.method == "POST" && req.path == "/excel/load") return handle_excel_load(req);
        if (req.method == "POST" && req.path == "/excel/query") return handle_excel_query(req);
        if (req.method == "POST" && req.path == "/excel/set") return handle_excel_set(req);
        if (req.method == "POST" && req.path == "/excel/close") return handle_excel_close(req);
        if (req.method == "POST" && req.path == "/apps/ui/get") return handle_ui_get(req);
        if (req.method == "POST" && req.path == "/apps/ui/save") return handle_ui_save(req);
        if (req.method == "GET" && req.path == "/apps") return handle_list(req);
        if (req.method == "POST" && req.path == "/apps") return handle_create(req);
        if (req.method == "POST" && req.path == "/apps/version") return handle_version_publish(req);
        if (req.method == "PUT" && req.path.rfind("/apps/", 0) == 0) return handle_update(req);
        if (req.method == "DELETE" && req.path.rfind("/apps/", 0) == 0) return handle_delete(req);
        if (req.method == "GET" && req.path == "/users") return handle_users_list(req);
        if (req.method == "POST" && req.path == "/users") return handle_users_upsert(req);
        HttpResponse resp;
        resp.status = 404;
        resp.body = "{\"error\":\"not found\"}";
        return resp;
    }

    bool require_json(const HttpRequest &req, HttpResponse &resp) {
        auto it = req.headers.find("Content-Type");
        if (it == req.headers.end()) return true; // allow silent fallback
        std::string v = it->second;
        for (auto &c : v) c = static_cast<char>(tolower(static_cast<unsigned char>(c)));
        if (v.find("application/json") == std::string::npos) {
            resp.status = 415;
            resp.body = "{\"error\":\"content-type must be application/json\"}";
            return false;
        }
        return true;
    }

    std::string bearer_token(const HttpRequest &req) {
        auto it = req.headers.find("Authorization");
        if (it == req.headers.end()) return "";
        const std::string &h = it->second;
        const std::string prefix = "Bearer ";
        if (h.rfind(prefix, 0) == 0) return h.substr(prefix.size());
        return "";
    }

    bool authenticate(const HttpRequest &req, UserRecord &user_out, HttpResponse &resp_out) {
        std::string token = bearer_token(req);
        std::string uname;
        if (token.empty() || !sessions_.verify(token, uname)) {
            resp_out.status = 401;
            resp_out.body = "{\"error\":\"unauthorized\"}";
            return false;
        }
        if (!db_.get_user(uname, user_out)) {
            resp_out.status = 403;
            resp_out.body = "{\"error\":\"user missing\"}";
            return false;
        }
        // force-admin from config
        if (cfg_.admins.count(uname)) user_out.role = Role::Admin;
        return true;
    }

    HttpResponse handle_login(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        std::string user = extract_json_string(req.body, "username");
        std::string pass = extract_json_string(req.body, "password");
        UserRecord u;
        if (!db_.get_user(user, u) || u.password != pass) {
            resp.status = 403;
            resp.body = "{\"error\":\"invalid credentials\"}";
            return resp;
        }
        if (cfg_.admins.count(user)) u.role = Role::Admin;
        std::string token = sessions_.login(user);
        resp.body = "{\"token\":\"" + token + "\"}";
        return resp;
    }

    HttpResponse handle_logout(const HttpRequest &req) {
        HttpResponse resp;
        std::string token = bearer_token(req);
        if (!token.empty()) sessions_.logout(token);
        resp.body = "{\"status\":\"ok\"}";
        return resp;
    }

    HttpResponse handle_health(const HttpRequest &) {
        HttpResponse resp;
        resp.body = "{\"status\":\"ok\"}";
        return resp;
    }

    HttpResponse handle_excel_load(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string owner = extract_json_string(req.body, "owner");
        if (owner.empty()) owner = caller.name;
        std::string app_name = extract_json_string(req.body, "name");
        int version = extract_json_int(req.body, "version", 0);
        if (app_name.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing app name\"}";
            return resp;
        }
        if (!is_safe_name(app_name) || !is_safe_name(owner)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid names\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(owner, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (!can_access(app, caller) && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        int ver = version > 0 ? version : app.latest_version;
        fs::path file_path = version_path(app.owner, app.name, ver) / (app.name + ".xlsx");
        if (!fs::exists(file_path)) {
            resp.status = 404;
            resp.body = "{\"error\":\"file not found\"}";
            return resp;
        }
        std::string token = bearer_token(req);
        if (token.empty()) {
            resp.status = 401;
            resp.body = "{\"error\":\"missing token\"}";
            return resp;
        }
        std::string err;
        if (!pool_.load_workbook(token, caller.name, file_path, err)) {
            resp.status = 503;
            resp.body = "{\"error\":\"" + json_escape(err) + "\"}";
            return resp;
        }
        resp.body = "{\"status\":\"loaded\",\"version\":" + std::to_string(ver) + "}";
        return resp;
    }

    HttpResponse handle_excel_query(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string sheet = extract_json_string(req.body, "sheet");
        std::string range = extract_json_string(req.body, "range");
        if (sheet.empty() || range.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"sheet and range required\"}";
            return resp;
        }
        std::string token = bearer_token(req);
        if (token.empty()) {
            resp.status = 401;
            resp.body = "{\"error\":\"missing token\"}";
            return resp;
        }
        std::string json_val;
        std::string err;
        if (!pool_.query_range(token, sheet, range, json_val, err)) {
            resp.status = 400;
            resp.body = "{\"error\":\"" + json_escape(err) + "\"}";
            return resp;
        }
        resp.body = "{\"value\":" + json_val + "}";
        return resp;
    }

    HttpResponse handle_excel_set(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string sheet = extract_json_string(req.body, "sheet");
        std::string range = extract_json_string(req.body, "range");
        if (sheet.empty() || range.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"sheet and range required\"}";
            return resp;
        }
        VARIANT val;
        VariantInit(&val);
        if (json_has_key(req.body, "value_bool")) {
            val.vt = VT_BOOL;
            val.boolVal = extract_json_bool(req.body, "value_bool", false) ? VARIANT_TRUE : VARIANT_FALSE;
        } else if (json_has_key(req.body, "value_number")) {
            val.vt = VT_R8;
            val.dblVal = extract_json_double(req.body, "value_number", 0.0);
        } else if (json_has_key(req.body, "value")) {
            std::string s = extract_json_string(req.body, "value");
            val.vt = VT_BSTR;
            val.bstrVal = SysAllocString(std::wstring(s.begin(), s.end()).c_str());
        } else {
            resp.status = 400;
            resp.body = "{\"error\":\"value required\"}";
            return resp;
        }
        std::string token = bearer_token(req);
        if (token.empty()) {
            VariantClear(&val);
            resp.status = 401;
            resp.body = "{\"error\":\"missing token\"}";
            return resp;
        }
        std::string err;
        bool ok = pool_.set_range_value(token, sheet, range, val, err);
        VariantClear(&val);
        if (!ok) {
            resp.status = 400;
            resp.body = "{\"error\":\"" + json_escape(err) + "\"}";
            return resp;
        }
        resp.body = "{\"status\":\"updated\"}";
        return resp;
    }

    HttpResponse handle_excel_close(const HttpRequest &req) {
        HttpResponse resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string token = bearer_token(req);
        if (token.empty()) {
            resp.status = 401;
            resp.body = "{\"error\":\"missing token\"}";
            return resp;
        }
        std::string err;
        if (!pool_.close_session(token, true, err)) {
            resp.status = 400;
            resp.body = "{\"error\":\"" + json_escape(err) + "\"}";
            return resp;
        }
        resp.body = "{\"status\":\"closed\",\"session\":\"restarted\"}";
        return resp;
    }

    HttpResponse handle_list(const HttpRequest &req) {
        HttpResponse resp;
        UserRecord u;
        if (!authenticate(req, u, resp)) return resp;
        resp.body = list_apps_json(u);
        return resp;
    }

    HttpResponse handle_create(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string app_name = extract_json_string(req.body, "name");
        std::string desc = extract_json_string(req.body, "description");
        std::string file_b64 = extract_json_string(req.body, "file_base64");
        std::string access_group = extract_json_string(req.body, "access_group");
        std::string image_b64 = extract_json_string(req.body, "image_base64");
        bool is_public = extract_json_bool(req.body, "public", false);
        if (app_name.empty() || file_b64.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing name or file_base64\"}";
            return resp;
        }
        if (!is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid app name\"}";
            return resp;
        }
        if (!access_group.empty() && !is_safe_name(access_group)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid access group\"}";
            return resp;
        }
        AppRecord existing;
        if (db_.get_app(caller.name, app_name, existing)) {
            resp.status = 409;
            resp.body = "{\"error\":\"app exists\"}";
            return resp;
        }
        fs::path ver_path = version_path(caller.name, app_name, 1);
        if (!ensure_dir(ver_path)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot create folder\"}";
            return resp;
        }
        std::string content = base64_decode(file_b64);
        std::ofstream file(ver_path / (app_name + ".xlsx"), std::ios::binary);
        file.write(content.data(), static_cast<std::streamsize>(content.size()));
        file.close();
        AppInfo info{app_name, 1, desc};
        write_metadata(ver_path, info);
        AppRecord rec{caller.name, app_name, 1, desc, is_public, access_group};
        db_.upsert_app(rec);
        if (!image_b64.empty()) {
            if (!save_app_image(caller.name, app_name, image_b64) ||
                !save_app_image_version(caller.name, app_name, 1, image_b64)) {
                resp.status = 500;
                resp.body = "{\"error\":\"cannot save image\"}";
                return resp;
            }
        }
        resp.body = "{\"status\":\"created\",\"version\":1}";
        return resp;
    }

    HttpResponse handle_update(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string app_name = req.path.substr(std::string("/apps/").size());
        if (app_name.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing app name\"}";
            return resp;
        }
        if (!is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid app name\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(caller.name, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (app.owner != caller.name && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        bool requested_new_version = extract_json_bool(req.body, "new_version", false);
        std::string desc = extract_json_string(req.body, "description");
        std::string file_b64 = extract_json_string(req.body, "file_base64");
        std::string access_group = extract_json_string(req.body, "access_group");
        bool public_flag = extract_json_bool(req.body, "public", app.public_access);
        std::string image_b64 = extract_json_string(req.body, "image_base64");
        if (requested_new_version) {
            resp.status = 400;
            resp.body = "{\"error\":\"use /apps/version to publish new versions\"}";
            return resp;
        }
        if (!access_group.empty() && !is_safe_name(access_group)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid access group\"}";
            return resp;
        }
        int ver = app.latest_version;
        fs::path target_path = version_path(app.owner, app.name, ver);
        if (!ensure_dir(target_path)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot create folder\"}";
            return resp;
        }
        if (!file_b64.empty()) {
            std::string content = base64_decode(file_b64);
            std::ofstream file(target_path / (app_name + ".xlsx"), std::ios::binary | std::ios::trunc);
            file.write(content.data(), static_cast<std::streamsize>(content.size()));
        }
        if (!desc.empty()) app.description = desc;
        if (!access_group.empty()) app.access_group = access_group;
        app.public_access = public_flag;
        app.latest_version = ver;
        AppInfo info{app_name, ver, app.description};
        write_metadata(target_path, info);
        db_.upsert_app(app);
        if (!image_b64.empty()) {
            if (!save_app_image(app.owner, app_name, image_b64) ||
                !save_app_image_version(app.owner, app_name, ver, image_b64)) {
                resp.status = 500;
                resp.body = "{\"error\":\"cannot save image\"}";
                return resp;
            }
        }
        resp.body = "{\"status\":\"updated\",\"version\":" + std::to_string(ver) + "}";
        return resp;
    }

    HttpResponse handle_version_publish(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string owner = extract_json_string(req.body, "owner");
        if (owner.empty()) owner = caller.name;
        std::string app_name = extract_json_string(req.body, "name");
        if (app_name.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing app name\"}";
            return resp;
        }
        if (!is_safe_name(owner) || !is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid names\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(owner, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (app.owner != caller.name && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        std::string desc = extract_json_string(req.body, "description");
        std::string file_b64 = extract_json_string(req.body, "file_base64");
        std::string image_b64 = extract_json_string(req.body, "image_base64");
        std::string schema_json = extract_json_string(req.body, "schema_json");
        std::string access_group = extract_json_string(req.body, "access_group");
        bool public_flag = extract_json_bool(req.body, "public", app.public_access);
        if (!access_group.empty() && !is_safe_name(access_group)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid access group\"}";
            return resp;
        }
        if (!schema_json.empty() && schema_json.size() > 256 * 1024) {
            resp.status = 400;
            resp.body = "{\"error\":\"schema too large\"}";
            return resp;
        }
        int prev_version = app.latest_version;
        int ver = prev_version + 1;
        fs::path target_path = version_path(app.owner, app.name, ver);
        if (!ensure_dir(target_path)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot create folder\"}";
            return resp;
        }
        fs::path target_file = target_path / (app_name + ".xlsx");
        bool workbook_ready = false;
        if (!file_b64.empty()) {
            std::string content = base64_decode(file_b64);
            std::ofstream file(target_file, std::ios::binary | std::ios::trunc);
            file.write(content.data(), static_cast<std::streamsize>(content.size()));
            file.close();
            workbook_ready = true;
        } else {
            fs::path prev_file = version_path(app.owner, app.name, prev_version) / (app_name + ".xlsx");
            std::error_code ec;
            if (fs::exists(prev_file)) {
                fs::copy_file(prev_file, target_file, fs::copy_options::overwrite_existing, ec);
                if (!ec) workbook_ready = true;
            }
        }
        if (!workbook_ready) {
            resp.status = 400;
            resp.body = "{\"error\":\"workbook missing\"}";
            return resp;
        }
        std::string schema_payload = schema_json;
        if (schema_payload.empty()) schema_payload = load_app_ui_version(app.owner, app_name, prev_version);
        if (schema_payload.empty()) schema_payload = load_app_ui(app.owner, app_name);
        if (schema_payload.empty()) schema_payload = "{\"components\":[]}";
        if (!save_app_ui_version(app.owner, app_name, ver, schema_payload) ||
            !save_app_ui(app.owner, app_name, schema_payload)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot save schema\"}";
            return resp;
        }
        bool had_prior_image = fs::exists(app_image_version_path(app.owner, app.name, prev_version)) ||
                               fs::exists(app_image_path(app.owner, app.name));
        bool image_ok = true;
        if (!image_b64.empty()) {
            image_ok = save_app_image(app.owner, app_name, image_b64) &&
                       save_app_image_version(app.owner, app_name, ver, image_b64);
        } else if (had_prior_image) {
            image_ok = copy_app_image_version(app.owner, app_name, prev_version, ver);
        }
        if (!image_ok) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot persist image\"}";
            return resp;
        }
        if (!desc.empty()) app.description = desc;
        if (!access_group.empty()) app.access_group = access_group;
        app.public_access = public_flag;
        app.latest_version = ver;
        AppInfo info{app_name, ver, app.description};
        write_metadata(target_path, info);
        db_.upsert_app(app);
        resp.body = "{\"status\":\"version_created\",\"version\":" + std::to_string(ver) + "}";
        return resp;
    }

    std::string list_apps_json(const UserRecord &viewer_in) {
        UserRecord viewer = viewer_in;
        std::vector<AppRecord> apps = db_.list_apps();
        std::ostringstream oss;
        oss << "[";
        bool first = true;
        for (auto &a : apps) {
            UserRecord owner_record;
            if (!db_.get_user(a.owner, owner_record)) continue;
            if (cfg_.admins.count(owner_record.name)) owner_record.role = Role::Admin;
            if (cfg_.admins.count(viewer.name)) viewer.role = Role::Admin;
            if (!can_access(a, viewer) && !is_admin(viewer, cfg_)) continue;
            if (!first) oss << ",";
            first = false;
            std::string image_data = load_app_image_base64_version(a.owner, a.name, a.latest_version);
            if (image_data.empty()) image_data = load_app_image_base64(a.owner, a.name);
            bool has_ui = fs::exists(app_ui_version_path(a.owner, a.name, a.latest_version)) || fs::exists(app_ui_path(a.owner, a.name));
            oss << "{\"owner\":\"" << json_escape(a.owner) << "\","
                << "\"name\":\"" << json_escape(a.name) << "\","
                << "\"latest_version\":" << a.latest_version << ","
                << "\"description\":\"" << json_escape(a.description) << "\","
                << "\"public\":" << (a.public_access ? "true" : "false") << ","
                << "\"access_group\":\"" << json_escape(a.access_group) << "\"," 
                << "\"has_ui\":" << (has_ui ? "true" : "false") << ","
                << "\"image_base64\":\"" << json_escape(image_data) << "\"}";
        }
        oss << "]";
        return oss.str();
    }

    HttpResponse handle_delete(const HttpRequest &req) {
        HttpResponse resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string app_name = req.path.substr(std::string("/apps/").size());
        if (app_name.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing app name\"}";
            return resp;
        }
        if (!is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid app name\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(caller.name, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (app.owner != caller.name && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        fs::path base = app_root() / app.owner / app_name;
        std::error_code ec;
        fs::remove_all(base, ec);
        db_.remove_app(app.owner, app_name);
        resp.body = "{\"status\":\"deleted\"}";
        return resp;
    }

    HttpResponse handle_ui_get(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string owner = extract_json_string(req.body, "owner");
        std::string app_name = extract_json_string(req.body, "name");
        if (owner.empty()) owner = caller.name;
        if (app_name.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing name\"}";
            return resp;
        }
        if (!is_safe_name(owner) || !is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid names\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(owner, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (!can_access(app, caller) && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        std::string schema = load_app_ui_version(owner, app_name, app.latest_version);
        if (schema.empty()) schema = load_app_ui(owner, app_name);
        if (schema.empty()) schema = "{\"components\":[]}";
        resp.body = "{\"owner\":\"" + json_escape(owner) + "\",\"name\":\"" + json_escape(app_name) + "\",\"schema\":" + schema + "}";
        return resp;
    }

    HttpResponse handle_ui_save(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        std::string owner = extract_json_string(req.body, "owner");
        std::string app_name = extract_json_string(req.body, "name");
        std::string schema_json = extract_json_string(req.body, "schema_json");
        if (owner.empty()) owner = caller.name;
        if (app_name.empty() || schema_json.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"missing fields\"}";
            return resp;
        }
        if (!is_safe_name(owner) || !is_safe_name(app_name)) {
            resp.status = 400;
            resp.body = "{\"error\":\"invalid names\"}";
            return resp;
        }
        if (schema_json.size() > 256 * 1024) {
            resp.status = 400;
            resp.body = "{\"error\":\"schema too large\"}";
            return resp;
        }
        AppRecord app;
        if (!db_.get_app(owner, app_name, app)) {
            resp.status = 404;
            resp.body = "{\"error\":\"not found\"}";
            return resp;
        }
        if (app.owner != caller.name && !is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"forbidden\"}";
            return resp;
        }
        if (!save_app_ui(owner, app_name, schema_json)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot save schema\"}";
            return resp;
        }
        if (!save_app_ui_version(owner, app_name, app.latest_version, schema_json)) {
            resp.status = 500;
            resp.body = "{\"error\":\"cannot save versioned schema\"}";
            return resp;
        }
        resp.body = "{\"status\":\"saved\"}";
        return resp;
    }

    HttpResponse handle_users_list(const HttpRequest &req) {
        HttpResponse resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        if (!is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"admin required\"}";
            return resp;
        }
        std::vector<UserRecord> users;
        db_.list_users(users);
        std::ostringstream oss;
        oss << "[";
        bool first = true;
        for (auto &u : users) {
            if (!first) oss << ",";
            first = false;
            oss << "{\"name\":\"" << json_escape(u.name) << "\",\"groups\":[";
            for (size_t i = 0; i < u.groups.size(); ++i) {
                if (i) oss << ",";
                oss << "\"" << json_escape(u.groups[i]) << "\"";
            }
            std::string role_str = "user";
            if (cfg_.admins.count(u.name)) role_str = "administrator";
            else if (u.role == Role::Developer) role_str = "developer";
            oss << "],\"role\":\"" << role_str << "\"}";
        }
        oss << "]";
        resp.body = oss.str();
        return resp;
    }

    HttpResponse handle_users_upsert(const HttpRequest &req) {
        HttpResponse resp;
        if (!require_json(req, resp)) return resp;
        UserRecord caller;
        if (!authenticate(req, caller, resp)) return resp;
        if (!is_admin(caller, cfg_)) {
            resp.status = 403;
            resp.body = "{\"error\":\"admin required\"}";
            return resp;
        }
        std::string name = extract_json_string(req.body, "username");
        std::string pass = extract_json_string(req.body, "password");
        std::string groups_csv = extract_json_string(req.body, "groups");
        std::string role_str = extract_json_string(req.body, "role");
        if (name.empty() || pass.empty()) {
            resp.status = 400;
            resp.body = "{\"error\":\"username and password required\"}";
            return resp;
        }
        Role new_role = Role::User;
        if (!role_str.empty()) {
            if (role_str == "user") new_role = Role::User;
            else if (role_str == "developer") new_role = Role::Developer;
            else {
                resp.status = 400;
                resp.body = "{\"error\":\"role must be user or developer\"}";
                return resp;
            }
        }
        if (cfg_.admins.count(name)) new_role = Role::Admin; // enforce config admin
        UserRecord existing;
        if (db_.get_user(name, existing)) {
            existing.password = pass;
            existing.groups = split_csv(groups_csv);
            if (cfg_.admins.count(name)) existing.role = Role::Admin; else existing.role = new_role;
        } else {
            existing = UserRecord{name, pass, split_csv(groups_csv), cfg_.admins.count(name) ? Role::Admin : new_role};
        }
        db_.upsert_user(existing);
        resp.body = "{\"status\":\"ok\"}";
        return resp;
    }

    const Config &cfg_;
    ExcelPool &pool_;
    Database &db_;
    SessionStore sessions_;
    bool running_ = false;
};

// -------------------- Entry --------------------
int main() {
    Config cfg = load_config("config.json");
    if (!ensure_dir(app_root())) {
        std::cerr << "Failed to create app root\n";
        return 1;
    }
    Database db("db.bin");
    if (!db.load()) {
        std::cerr << "Failed to load db.bin\n";
        return 1;
    }
    // Seed or sync config-defined admins
    for (const auto &admin : cfg.admins) {
        UserRecord u;
        std::string pass = "admin";
        auto it = cfg.users.find(admin);
        if (it != cfg.users.end()) pass = it->second;
        if (db.get_user(admin, u)) {
            u.role = Role::Admin;
            if (!pass.empty()) u.password = pass;
            db.upsert_user(u);
        } else {
            UserRecord nu{admin, pass, std::vector<std::string>{"admin"}, Role::Admin};
            db.upsert_user(nu);
        }
    }
    ExcelPool pool;
    if (!pool.init(cfg.excel_instances)) {
        std::cerr << "Excel pool initialization failed\n";
        return 1;
    }
    Server srv(cfg, pool, db);
    srv.start();
    pool.shutdown();
    return 0;
}
