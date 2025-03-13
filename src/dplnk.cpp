#include "dplnk.h"

#ifdef _WIN32 // Windows
#include "subsystems/windows.h"
#endif

void dplnk::dplnk(const std::string& path, dplnk::options options) {
#ifdef _WIN32 // Windows
    const std::wstring wprotocol(options.protocol.begin(), options.protocol.end());

    WinReg::RegKey protocolkey{ HKEY_CLASSES_ROOT, wprotocol };
    protocolkey.setStringValue(L"", L"URL: " + wprotocol + L" Protocol");
    protocolkey.setStringValue(L"URL Protocol", L"");

    WinReg::RegKey iconkey{ HKEY_CLASSES_ROOT, wprotocol + L"\\DefaultIcon" };
    iconkey.setStringValue(L"", L"C:\\Windows\\System32\\url.dll,0");

    WinReg::RegKey cmdkey{ HKEY_CLASSES_ROOT, wprotocol + L"\\shell\\open\\command" };

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