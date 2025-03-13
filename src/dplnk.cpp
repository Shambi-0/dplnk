#include "dplnk.h"

#ifdef _WIN32 // Windows
#include "subsystems/windows.h"
#endif

void dplnk::dplnk(const std::string& path, dplnk::options options) {
#ifdef _WIN32 // Windows
    const std::wstring wprotocol(options.protocol.begin(), options.protocol.end());
    WinReg::RegKey cmdkey{ HKEY_CURRENT_USER, wprotocol + L"\\shell\\open\\command" };

    const std::wstring wpath(path.begin(), path.end());
    cmdkey.setStringValue(L"", L"\"" + wpath + L"\" %1");

    if (options.d.has_value()) {
        for (const auto& [key, value] : *options.d) {
            const std::wstring wkey(key.begin(), key.end());
            const std::wstring wvalue(value.begin(), value.end());

            cmdkey.setStringValue(wkey, wvalue);
        }
    }
#else
    throw std::runtime_error("Unsupported platform!");
#endif
}