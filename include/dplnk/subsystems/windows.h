#pragma once

#include <Windows.h>
#include <crtdbg.h>

#include <limits>
#include <memory>
#include <stdexcept>
#include <string>
#include <system_error>
#include <utility>
#include <variant>
#include <vector>

namespace WinReg {
	class RegException;
	class RegResult;

	template<typename T>
	class RegExpected;

	class RegKey {
	public:
		// Initializes the object with an empty key handle
		RegKey() noexcept = default;
		
		// Take ownership of the input key handle
		explicit RegKey(HKEY hKey) noexcept;

		/**
		* Open the given registry key if it exists, otherwise creates a new key.
		* Uses the default `KEY_READ|KEY_WRITE|KEY_WOW64_64KEY` acccess.
		* For finer grained control, call the `Create()` method overloads.
		* Throws `RegException` on failure.
		*/
		RegKey(HKEY hKeyParent, const std::wstring& subKey);

		/**
		* Open the registry key if it exists, otherwise creates a new key.
		* Allow the caller to specify the desired access to the key.
		* (e.g. `KEY_READ|KEY_WOW64_64KEY` for read-only access).
		* For finer grained control, call the `Create()` method overloads.
		* Throws `RegException` on failure.
		*/
		RegKey(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess);

		/**
		* Take ownership of the input key handle.
		* The input key handle wrapper is reset to an empty state.
		*/
		RegKey(RegKey&& other) noexcept;

		/**
		* Move-assign from the input key handle.
		* Properly check against self-move-assign (which is safe & does nothing).
		*/
		RegKey& operator=(RegKey&& other) noexcept;

		// Ban copy
		RegKey(const RegKey&) = delete;
		RegKey& operator=(const RegKey&) = delete;

		// Safely close the wrapped key handle (if any)
		~RegKey() noexcept;

		// Access the wrapped raw `HKEY` handle
		[[nodiscard]] HKEY get() const noexcept;

		// Is the wrapped `HKEY` handle valid?
		[[nodiscard]] bool isValid() const noexcept;

		// Same as `isValid()`, but allows for a shorter syntax
		[[nodiscard]] explicit operator bool() const noexcept;

		// Is the wrapped handle a predefined handle (e.g. `HKEY_CURRENT_USER`)?
		[[nodiscard]] bool isPredefined() const noexcept;

		/**
		* @brief Closes the current `HKEY` handle.
		* @note If there isn't a valid handle, does nothing.
		* @note This method doesn't close predefined `HKEY` handles (e.g. `HKEY_CURRENT_USER`).
		*/
		void close() noexcept;

		/**
		* @details Transfer ownership of the current `HKEY` to the caller.
		* @warning The caller is responsible for closing the key handle!
		*/
		[[nodiscard]] HKEY detach() noexcept;

		/**
		* Take ownership of the input `HKEY` handle.
		* Safely close any previously open handle.
		* Input key handle can be a `nullptr`.
		*/
		void attach(HKEY hKey) noexcept;

		/**
		* Non-throwing swap;
		* @note There is also a non-member swap overload.
		*/
		void swapWith(RegKey& other) noexcept;

		/**
		* By default, a 32-bit application running on a 64-bit version of Windows still accesses the 32-bit view of the registry,
		* and a 64-bit application accesses the 64-bit view of the registry.
		* 
		* Using this `KEY_WOW64_64KEY` flag, both 32-bit or 64-bit applications access the 64-bit registry view.
		* 
		* MSDN documentation:
		* @link https://docs.microsoft.com/en-us/windows/win32/winprog64/accessing-an-alternate-registry-view
		* 
		* If you want to use the default Windows API behavior, don't use OR (|) the `KEY_WOW64_64KEY` flag
		* when specifying the desired access (e.g. just pass `KEY_READ|KEY_WRITE` as the desired access parameter)
		*/

		// Wrapper around `RegCreateKeyEx()`, that allows you to specify the desired access.
		void create(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY);

		// Wrapper around `RegCreateKeyEx()`.
		void create(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess, DWORD options, SECURITY_ATTRIBUTES* securityAttributes, DWORD* disposition);

		// Wrapper around `RegCreateKeyEx()`.
		void open(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY);

		// Wrapper around `RegCreateKeyEx()`, that allows you to specify the desired access.
		[[nodiscard]] RegResult tryCreate(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY);

		// Wrapper around `RegCreateKeyEx()`.
		[[nodiscard]] RegResult tryCreate(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess, DWORD options, SECURITY_ATTRIBUTES* securityAttributes, DWORD* disposition);

		// Wrapper around `RegCreateKeyEx()`.
		[[nodiscard]] RegResult tryOpen(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY);

		// Registry value setters
		void setDwordValue(const std::wstring& valueName, DWORD data);
		void setQwordValue(const std::wstring& valueName, const ULONGLONG& data);
		void setStringValue(const std::wstring& valueName, const std::wstring& data);
		void setExpandStringValue(const std::wstring& valueName, const std::wstring& data);
		void setMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data);
		void setBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data);
		void setBinaryValue(const std::wstring& valueName, const void* data, DWORD dataSize);

		/**
		* Registry value setters which return a `RegResult`, instead of throwing a `RegException`.
		*/
		[[nodiscard]] RegResult trySetDwordValue(const std::wstring& valueName, DWORD data) noexcept;
		[[nodiscard]] RegResult trySetQwordValue(const std::wstring& valueName, const ULONGLONG& data) noexcept;

		/**
		* @note The `trySetStringValue` method cannot be marked as `noexcept`,
		* because internally the method "may" throw an exception if the input string is too big
		* (`size_t` value overflowing a `DWORD`)
		*/
		[[nodiscard]] RegResult trySetStringValue(const std::wstring& valueName, const std::wstring& data);
		
		/**
		* @note The `trySetExpandStringValue` method cannot be marked as `noexcept`,
		* because internally the method "may" throw an exception if the input string is too big
		* (`size_t` value overflowing a `DWORD`)
		*/
		[[nodiscard]] RegResult trySetExpandStringValue(const std::wstring& valueName, const std::wstring& data);

		/**
		* @note The `trySetMultiStringValue` method cannot be marked as `noexcept`,
		* because internally the method "dynamically allocates memory" for creating the multi-string that will be stored in the registry.
		*/
		[[nodiscard]] RegResult trySetMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data);

		/**
		* @note This overload of the `trySetBinaryValue()` method cannot be marked `noexcept`,
		* because internally the method "may" throw an exception if the input vector is too large
		* (`vector::size size_t` value overflowing a `DWORD`)
		*/
		[[nodiscard]] RegResult trySetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data);

		[[nodiscard]] RegResult trySetBinaryValue(const std::wstring& valueName, const void* data, DWORD dataSize) noexcept;

		// Registry value getters
		[[nodiscard]] DWORD getDwordValue(const std::wstring& valueName) const;
		[[nodiscard]] ULONGLONG getQwordValue(const std::wstring& valueName) const;
		[[nodiscard]] std::wstring getStringValue(const std::wstring& valueName) const;

		enum class ExpandStringOption {
			DontExpand,
			Expand
		};

		[[nodiscard]] std::wstring getExpandStringValue(const std::wstring& valueName, ExpandStringOption expandOption = ExpandStringOption::DontExpand) const;

		[[nodiscard]] std::vector<std::wstring> getMultiStringValue(const std::wstring& valueName) const;
		[[nodiscard]] std::vector<BYTE> getBinaryValue(const std::wstring& valueName) const;

		/**
		* Registry value getters which return a `RegExpected<T>`, instead of throwing a `RegException`.
		*/
		[[nodiscard]] RegExpected<DWORD> tryGetDwordValue(const std::wstring& valueName) const;
		[[nodiscard]] RegExpected<ULONGLONG> tryGetQwordValue(const std::wstring& valueName) const;
		[[nodiscard]] RegExpected<std::wstring> tryGetStringValue(const std::wstring& valueName) const;

		[[nodiscard]] RegExpected<std::wstring> tryGetExpandStringValue(const std::wstring& valueName, ExpandStringOption expandOption = ExpandStringOption::DontExpand) const;

		[[nodiscard]] RegExpected<std::vector<std::wstring>> tryGetMultiStringValue(const std::wstring& valueName) const;
		[[nodiscard]] RegExpected<std::vector<BYTE>> tryGetBinaryValue(const std::wstring& valueName) const;

		// Query operations

		// Information about a registry key (retrieved by `QueryInfoKey`)
		struct InfoKey {
			DWORD NumberOfSubKeys;
			DWORD NumberOfValues;
			FILETIME LastWriteTime;

			// Clear the structure fields
			InfoKey() noexcept : NumberOfSubKeys{ 0 }, NumberOfValues{ 0 } {
				LastWriteTime.dwHighDateTime = LastWriteTime.dwLowDateTime = 0;
			};

			InfoKey(DWORD numberOfSubKeys, DWORD numberOfValues, FILETIME lastWriteTime) noexcept :
				NumberOfSubKeys{ numberOfSubKeys }, NumberOfValues{ numberOfValues }, LastWriteTime{ lastWriteTime } {
			};
		};

		// Retrieve information about the registry key
		[[nodiscard]] InfoKey queryInfoKey() const;

		// Return the `DWORD` type ID for the input registry value
		[[nodiscard]] DWORD queryValueType(const std::wstring& valueName) const;

		enum class KeyReflection {
			ReflectionEnabled,
			ReflectionDisabled
		};

		// Determines whether reflection has been disabled for the specified key
		[[nodiscard]] KeyReflection queryReflectionKey() const;

		// Enumerate the subkeys of the registry key, using `RegEnumKeyEx`.
		[[nodiscard]] std::vector<std::wstring> enumSubKeys() const;

		// Enumerate the values under the registry key, using `RegEnumValue`.
		[[nodiscard]] std::vector<std::pair<std::wstring, DWORD>> enumValues() const;

		// Check if the current key contains a specific value.
		[[nodiscard]] bool hasValue(const std::wstring& valueName) const;

		// Check if the current key contains the specified sub-key.
		[[nodiscard]] bool hasSubKey(const std::wstring& subKey) const;

		/**
		* Query operations which return a `RegExpected<T>`, instead of throwing a `RegException`.
		*/

		// Retrieve information about the registry key
		[[nodiscard]] RegExpected<InfoKey> tryQueryInfoKey() const;

		// Return the `DWORD` type ID for the input registry value
		[[nodiscard]] RegExpected<DWORD> tryQueryValueType(const std::wstring& valueName) const;

		// Determines whether reflection has been disabled for the specified key
		[[nodiscard]] RegExpected<KeyReflection> tryQueryReflectionKey() const;

		// Enumerate the subkeys of the registry key, using `RegEnumKeyEx`.
		[[nodiscard]] RegExpected<std::vector<std::wstring>> tryEnumSubKeys() const;

		// Enumerate the values under the registry key, using `RegEnumValue`.
		[[nodiscard]] RegExpected<std::vector<std::pair<std::wstring, DWORD>>> tryEnumValues() const;

		// Check if the current key contains a specific value.
		[[nodiscard]] RegExpected<bool> tryHasValue(const std::wstring& valueName) const;

		// Check if the current key contains the specified sub-key.
		[[nodiscard]] RegExpected<bool> tryHasSubKey(const std::wstring& subKey) const;

		// Miscellaneous registry API wrappers
		void deleteValue(const std::wstring& valueName);
		void deleteKey(const std::wstring& subKey, REGSAM desiredAccess);
		void deleteTree(const std::wstring& subKey);
		void copyTree(const std::wstring& srcSubKey, const RegKey& dstKey);
		void flushKey();
		void loadKey(const std::wstring& subKey, const std::wstring& filename);
		void saveKey(const std::wstring& filename, SECURITY_ATTRIBUTES* securityAttributes) const;
		void enableReflectionKey();
		void disableReflectionKey();
		void connectRegistry(const std::wstring& machineName, HKEY hKeyPredefined);

		/**
		* Miscellaneous registry API wrappers which return a `RegResult`, instead of throwing a `RegException`.
		*/
		[[nodiscard]] RegResult tryDeleteValue(const std::wstring& valueName) noexcept;
		[[nodiscard]] RegResult tryDeleteKey(const std::wstring& subKey, REGSAM desiredAccess) noexcept;
		[[nodiscard]] RegResult tryDeleteTree(const std::wstring& subKey) noexcept;
		[[nodiscard]] RegResult tryCopyTree(const std::wstring& srcSubKey, const RegKey& dstKey) noexcept;
		[[nodiscard]] RegResult tryFlushKey() noexcept;
		[[nodiscard]] RegResult tryLoadKey(const std::wstring& subKey, const std::wstring& filename) noexcept;
		[[nodiscard]] RegResult trySaveKey(const std::wstring& filename, SECURITY_ATTRIBUTES* securityAttributes) const noexcept;
		[[nodiscard]] RegResult tryEnableReflectionKey() noexcept;
		[[nodiscard]] RegResult tryDisableReflectionKey() noexcept;
		[[nodiscard]] RegResult tryConnectRegistry(const std::wstring& machineName, HKEY hKeyPredefined) noexcept;

		// Return a string representation of the Windows registry types
		[[nodiscard]] static std::wstring regTypeToString(DWORD regType);

	private:
		// The wrapped registry `HKEY` handle
		HKEY m_hKey{ nullptr };
	}; // class RegKey

	/**
	* An exception representing an error with the registry operations.
	*/
	class RegException : public std::system_error {
	public:
		RegException(LSTATUS errorCode, const char* message);
		RegException(LSTATUS errorCode, const std::string& message);
	};

	/**
	* A tiny wrapper around the `LSTATUS` return codes used by the Windows Registry API.
	*/
	class RegResult {
	public:
		// Initialize with the success code `ERROR_SUCCESS`.
		RegResult() noexcept = default;

		// Initialize with a specific Windows Registry API `LSTATUS` return code.
		explicit RegResult(LSTATUS result) noexcept;

		// Check if the wrapped code is a success code (ERROR_SUCCESS)?
		[[nodiscard]] bool isOk() const noexcept;

		// Check if the wrapped error code is a failure code?
		[[nodiscard]] bool failed() const noexcept;

		// Check if the wrapped error code is a success code?
		[[nodiscard]] explicit operator bool() const noexcept;

		// Get the wrapped Win32 return code
		[[nodiscard]] LSTATUS code() const noexcept;

		// Return the system error message associated with the current error code
		[[nodiscard]] std::wstring message() const;

		// Return the system error message associated with the current error code, using the given locale
		[[nodiscard]] std::wstring message(DWORD locale) const;

	private:
		// Error code returned by the Windows Registry C API, initialized with the success code.
		LSTATUS m_result{ ERROR_SUCCESS };
	}; // class RegResult

	/**
	* A class template that stores the value of a type `T` (e.g. `DWORD`, `std::wstring`)
	* on success, or a `RegResult` on failure.
	* 
	* Used as the return value of some registry `RegKey::tryGetXxxValue()` methods
	* as an alternative to throwing exceptions.
	*/
	template<typename T>
	class RegExpected {
	public:
		// Initialize the object with an error code
		explicit RegExpected(const RegResult& errorCode) noexcept;

		// Initialize the object with a value of type `T` (on success)
		explicit RegExpected(const T& value);

		// Initialize the object with a value of type `T` (on success), optimized for move semantics.
		explicit RegExpected(T&& value);

		// Check if this object contains a valid value of type `T`?
		[[nodiscard]] bool isValid() const noexcept;

		// Check if this object contains a valid value of type `T`?
		[[nodiscard]] explicit operator bool() const noexcept;

		/**
		* Access the value (if the object contains a valid value)
		* But throws an exception if the object is in an invalid state.
		*/
		[[nodiscard]] const T& getValue() const;

		/**
		* Access the error code (if the object contains an error status)
		* But throws an exception if the object is in a valid state.
		*/
		[[nodiscard]] RegResult getError() const;

	private:
		// Stores a value of type `T` (on success), or a `RegResult` (on failure)
		std::variant<RegResult, T> m_var;
	}; // class RegExpected

	// Overloads of relational comparison operators for `RegKey`
	inline bool operator==(const RegKey& a, const RegKey& b) noexcept {
		return a.get() == b.get();
	}

	inline bool operator!=(const RegKey& a, const RegKey& b) noexcept {
		return a.get() != b.get();
	}

	inline bool operator<(const RegKey& a, const RegKey& b) noexcept {
		return a.get() < b.get();
	}

	inline bool operator<=(const RegKey& a, const RegKey& b) noexcept {
		return a.get() <= b.get();
	}

	inline bool operator>(const RegKey& a, const RegKey& b) noexcept {
		return a.get() > b.get();
	}

	inline bool operator>=(const RegKey& a, const RegKey& b) noexcept {
		return a.get() >= b.get();
	}

	// Private helper classes & functions
	namespace Details {
		/**
		 * Simple scope-based RAII wrapper that automatically invokes `LocalFree` in it's destructor.
		 */
		template<typename T>
		class ScopedLocalFree {
		public:
			// Initialize wrapped pointer to `nullptr`
			ScopedLocalFree() noexcept = default;

			// Automatically & safely free the wrapped pointer
			~ScopedLocalFree() noexcept {
				this->free();
			}

			// Ban copy & move operations
			ScopedLocalFree(const ScopedLocalFree&) = delete;
			ScopedLocalFree(ScopedLocalFree&&) = delete;
			ScopedLocalFree& operator=(const ScopedLocalFree&) = delete;
			ScopedLocalFree& operator=(ScopedLocalFree&&) = delete;

			// Read-only access to the wrapped pointer
			[[nodiscard]] T* get() const noexcept {
				return m_ptr;
			}

			// Writable access to the wrapped pointer
			[[nodiscard]] T** addressOf() noexcept {
				return &m_ptr;
			}

			// Explicit pointer conversion to bool
			explicit operator bool() const noexcept {
				return m_ptr != nullptr;
			}

			// Safely release the wrapped pointer
			void free() noexcept {
				if (m_ptr != nullptr) {
					::LocalFree(m_ptr);
					m_ptr = nullptr;
				}
			}

		private:
			T* m_ptr{ nullptr };
		}; // class ScopedLocalFree

		/**
		 * @brief Helper function to build a multi-string from a `std::vector<std::wstring>`.
		 * 
		 * A multi-string is a sequence of null-terminated strings, that terminates with an additional null character.
		 * Basically, considered as a whole, the sequence is terminated by two null characters.
		 * 
		 * E.g.:
		 * 		Hello\0World\0\0
		 */
		[[nodiscard]] inline std::vector<wchar_t> buildMultiString(const std::vector<std::wstring>& data) {
			// Special case of an empty multi-string
			if (data.empty()) {
				// Build a vector containing two null characters
				return std::vector<wchar_t>(2, L'\0');
			}

			// Calculate the total length in `wchar_t`s of the multi-string
			size_t n = 0;
			for (const auto& s : data) {
				// Add one to the current string length for the null terminator
				n += s.length() + 1;
			}

			// Add one more for the final null terminator
			n++;

			// Allocate a buffer to store the multi-string
			std::vector<wchar_t> multiString;

			// Reserve room in the buffer to speed up the following insertion loop
			multiString.reserve(n);

			// Copy the single strings into the multi-string buffer
			for (const auto& s : data) {
				if (!s.empty()) {
					// Copy the current string's contents
					multiString.insert(multiString.end(), s.begin(), s.end());
				}

				// Don't forget to add the null terminator (or just insert L'\0' for empty strings)
				multiString.emplace_back(L'\0');
			}

			// Add the final null terminator
			multiString.emplace_back(L'\0');

			return multiString;
		}

		// Return true if the `std::wchar_t` sequence stored in `data` terminates, with two nullptr (L'\0') `std::wchar_t`s.
		[[nodiscard]] inline bool isDoubleNullTerminated(const std::vector<wchar_t>& data) {
			// First check that there is enough room for at least two null.
			if (data.size() < 2) return false;

			// Check the last sequence terminates with two nullptr (L'\0', L'\0')
			const size_t n = data.size() - 1;
			return (data[n] == L'\0' && data[n - 1] == L'\0');
		}

		// Given a sequence of `std::wchar_t`s representing a double-null-terminated multi-string, returns a `std::vector<std::wstring>`.
		[[nodiscard]] inline std::vector<std::wstring> parseMultiString(const std::vector<wchar_t>& data) {
			// Make sure that there are terminating L'\0's at the end of the sequence
			if (!isDoubleNullTerminated(data)) {
				throw RegException{ ERROR_INVALID_DATA, "Not a double-nullptr terminated string." };
			}

			// Parse the double-null-terminated string into a `std::vector<std::wstring>`, which will be returned to the caller
			std::vector<std::wstring> result;

			const wchar_t* ptr = data.data();
			const wchar_t* const ptrEnd = data.data() + data.size() - 1;

			while (ptr < ptrEnd) {
				// Current string is nullptr-terminated, so get it's length calling `wcslen`
				const size_t len = wcslen(ptr);

				// Append the current string to the result vector
				if (len > 0) {
					result.emplace_back(ptr, len);
				}
				else {
					// Empty string, just append an empty string
					result.emplace_back(std::wstring{});
				}

				// Move to the next string, skipping the null terminator
				ptr += len + 1;
			}

			return result;
		}

		// Builds a `RegExpected<T>` object that stores an error code
		template<typename T>
		[[nodiscard]] inline RegExpected<T> makeRegExpectedWithError(const LSTATUS errorCode) noexcept {
			return RegExpected<T>{ RegResult{ errorCode } };
		}

		// Returns true if casting a `size_t` value to a `DWORD` is safe, (e.g. there is no overflow); false otherwise.
		[[nodiscard]] inline bool sizeToDwordCastIsSafe([[maybe_unused]] const size_t size) noexcept {
#ifdef _WIN64
			/**
			 * In 64-bit builds, `DWORD` is an unsigned 32-bit integer, while `size_t` is an unsigned 64-bit integer.
			 * So we need to pay attention to the conversion from `size_t` to `DWORD`.
			 */

			/**
			 * Pre-compute at compile-time the maximum value that can be stored by a `DWORD`.
			 * @note That this value is stored in a `size_t` for the proper comparison with the `size` parameter.
			 */
			constexpr size_t kMaxDwordValue = static_cast<size_t>((std::numeric_limits<DWORD>::max)());

			// Check against overflow
			if (size > kMaxDwordValue) {
				// Overflow from `size_t` to `DWORD`
				return false;
			}

			// The conversion is safe on 64-bit Windows
			return true;
#else
			static_assert(sizeof(size_t) == sizeof(DWORD)); // Both are 32-bit unsigned integers on 32-bit x86
			return true;
#endif // _WIN64
		}

		/**
		 * Safely cast a `size_t` value (usually from the STL) to a `DWORD` (usually for Win32 API Calls).
		 * In case of overflow, throws an exception of type `std::overflow_error`.
		 */
		[[nodiscard]] inline DWORD safeCastSizeToDword(const size_t size) {
#ifdef _WIN64
			/**
			 * In 64-bit builds, `DWORD` is an unsigned 32-bit integer, while `size_t` is an unsigned 64-bit integer.
			 * So we need to pay attention to the conversion from `size_t` to `DWORD`.
			 */

			// Check against overflow
			if (!sizeToDwordCastIsSafe(size)) {
				throw std::overflow_error("Overflow when casting 'size_t' to 'DWORD'.");
			}

			return static_cast<DWORD>(size);
#else
			/**
			 * In 32-bit builds with Microsoft Visual C++, a `size_t` is an unsigned 32-bit value,
			 * just like a `DWORD`. So, we can optimize this case out for 32-bit builds.
			 */
			_ASSERTE(sizeToDwordCastIsSafe(size)); // double-check just in debug builds
			static_assert(); // Both 32-bit unsigned integers on
#endif // _WIN64
		}
	} // namespace Details

	// RegKey inline methods
	inline RegKey::RegKey(const HKEY hKey) noexcept : m_hKey{ hKey } {}

	inline RegKey::RegKey(const HKEY hKeyParent, const std::wstring& subKey) {
		create(hKeyParent, subKey);
	}

	inline RegKey::RegKey(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess) {
		create(hKeyParent, subKey, desiredAccess);
	}

	inline RegKey::RegKey(RegKey&& other) noexcept : m_hKey{ other.m_hKey } {
		// "other" doesn't own the handle anymore
		other.m_hKey = nullptr;
	}

	inline RegKey& RegKey::operator=(RegKey&& other) noexcept {
		// Prevent self-move-assign
		if ((this != &other) && (m_hKey != other.m_hKey)) {
			// Close current
			close();

			// Move from the other (i.e. take ownership of other's raw handle)
			m_hKey = other.m_hKey;
			other.m_hKey = nullptr;
		}

		return *this;
	}

	inline RegKey::~RegKey() noexcept {
		// Release the owned handle (if any)
		close();
	}

	inline HKEY RegKey::get() const noexcept {
		return m_hKey;
	}

	inline void RegKey::close() noexcept {
		if (isValid()) {
			// Do not call `RegCloseKey` on predefined keys
			if (!isPredefined()) {
				::RegCloseKey(m_hKey);
			}

			// Avoid dangling references
			m_hKey = nullptr;
		}
	}

	inline bool RegKey::isValid() const noexcept {
		return m_hKey != nullptr;
	}

	inline RegKey::operator bool() const noexcept {
		return isValid();
	}

	inline bool RegKey::isPredefined() const noexcept {
		/**
		 * Predefined keys
		 * @link https://msdn.microsoft.com/en-us/library/windows/desktop/ms724836(v=vs.85).aspx
		 */
		if (   (m_hKey == HKEY_CURRENT_USER)
			|| (m_hKey == HKEY_LOCAL_MACHINE)
			|| (m_hKey == HKEY_CLASSES_ROOT)
			|| (m_hKey == HKEY_CURRENT_CONFIG)
			|| (m_hKey == HKEY_CURRENT_USER_LOCAL_SETTINGS)
			|| (m_hKey == HKEY_PERFORMANCE_DATA)
			|| (m_hKey == HKEY_PERFORMANCE_NLSTEXT)
			|| (m_hKey == HKEY_PERFORMANCE_TEXT)
			|| (m_hKey == HKEY_USERS))
		{
			return true;
		}

		return false;
	}

	inline HKEY RegKey::detach() noexcept {
		HKEY hKey = m_hKey;

		// We don't own the `HKEY` handle anymore
		m_hKey = nullptr;

		// Transfer ownership to the caller
		return hKey;
	}

	inline void RegKey::attach(const HKEY hKey) noexcept {
		// Prevent self-attach
		if (m_hKey != hKey) {
			// Close the current handle
			close();

			// Take ownership of the input handle
			m_hKey = hKey;
		}
	}

	inline void RegKey::swapWith(RegKey& other) noexcept {
		// Enable ADL (not necessary in this case, but good practice)
		// Swap the raw handle members
		std::swap(m_hKey, other.m_hKey);
	}

	inline void swap(RegKey& a, RegKey& b) noexcept {
		a.swapWith(b);
	}

	inline void RegKey::create(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess) {
		constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

		create(hKeyParent, subKey, desiredAccess, kDefaultOptions,
			nullptr, // No security attributes
			nullptr // No disposition
		);
	}

	inline void RegKey::create(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess, const DWORD options, SECURITY_ATTRIBUTES* const securityAttributes, DWORD* const disposition) {
		HKEY hKey = nullptr;
		LSTATUS errorCode = ::RegCreateKeyExW(
			hKeyParent,
			subKey.c_str(),
			0, // Reserved, must be zero
			REG_NONE, // User-defined class type parameter not supported
			options,
			desiredAccess,
			securityAttributes,
			&hKey,
			disposition
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "RegCreateKeyExW failed." };
		}

		// Safely close any previously opened key
		close();

		// Take ownership of the newly created key
		m_hKey = hKey;
	}

	inline void RegKey::open(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess) {
		HKEY hKey = nullptr;
		LSTATUS errorCode = ::RegOpenKeyExW(
			hKeyParent,
			subKey.c_str(),
			REG_NONE, // default options
			desiredAccess,
			&hKey
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "RegOpenKeyExW failed." };
		}

		// Safely close any previously opened key
		close();

		// Take ownership of the newly created key
		m_hKey = hKey;
	}

	inline RegResult RegKey::tryCreate(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY) {
		constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

		return tryCreate(hKeyParent, subKey, desiredAccess, kDefaultOptions,
			nullptr, // No security attributes
			nullptr // No disposition
		);
	}

	inline RegResult RegKey::tryCreate(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess, const DWORD options, SECURITY_ATTRIBUTES* const securityAttributes, DWORD* const disposition) noexcept {
		HKEY hKey = nullptr;
		RegResult errorCode{ ::RegCreateKeyExW(
			hKeyParent,
			subKey.c_str(),
			0, // Reserved, must be zero
			REG_NONE, // User-defined class type parameter not supported
			options,
			desiredAccess,
			securityAttributes,
			&hKey,
			disposition
		) };

		if (errorCode.failed()) {
			return errorCode;
		}

		// Safely close any previously opened key
		close();

		// Take ownership of the newly created key
		m_hKey = hKey;

		_ASSERTE(errorCode.isOk());
		return errorCode;
	}

	inline RegResult RegKey::tryOpen(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY) noexcept {
		HKEY hKey = nullptr;
		RegResult errorCode{ ::RegOpenKeyExW(
			hKeyParent,
			subKey.c_str(),
			REG_NONE, // default options
			desiredAccess,
			&hKey
		) };

		if (errorCode.failed()) {
			return errorCode;
		}

		// Safely close any previously opened key
		close();

		// Take ownership of the newly created key
		m_hKey = hKey;

		_ASSERTE(errorCode.isOk());
		return errorCode;
	}

	inline void RegKey::setDwordValue(const std::wstring& valueName, const DWORD data) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_DWORD,
			reinterpret_cast<const BYTE*>(&data),
			sizeof(data)
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write DWORD value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setQwordValue(const std::wstring& valueName, const ULONGLONG& data) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_QWORD,
			reinterpret_cast<const BYTE*>(&data),
			sizeof(data)
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write QWORD value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setStringValue(const std::wstring& valueName, const std::wstring& data) {
		_ASSERTE(isValid());

		// String size including the terminating null character, in bytes
		const DWORD dataSize = Details::safeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_SZ,
			reinterpret_cast<const BYTE*>(data.c_str()),
			dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write string value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setExpandStringValue(const std::wstring& valueName, const std::wstring& data) {
		_ASSERTE(isValid());

		// String size including the terminating null character, in bytes
		const DWORD dataSize = Details::safeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_EXPAND_SZ,
			reinterpret_cast<const BYTE*>(data.c_str()),
			dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write expand string value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data) {
		_ASSERTE(isValid());

		// First, we have to build a double-null-terminated multi-string from the input data
		const std::vector<wchar_t> multiString = Details::buildMultiString(data);

		// Total size, in bytes, of the whole multi-string sequence
		const DWORD dataSize = Details::safeCastSizeToDword(multiString.size() * sizeof(wchar_t));

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_MULTI_SZ,
			reinterpret_cast<const BYTE*>(multiString.data()),
			dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write multi-string value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data) {
		_ASSERTE(isValid());

		// Total size, in bytes, of the binary data
		const DWORD dataSize = Details::safeCastSizeToDword(data.size());

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_BINARY,
			data.data(),
			dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write binary value: RegSetValueExW failed." };
		}
	}

	inline void RegKey::setBinaryValue(const std::wstring& valueName, const void* const data, const DWORD dataSize) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_BINARY,
			static_cast<const BYTE*>(data),
			dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot write binary value: RegSetValueExW failed." };
		}
	}

	inline RegResult RegKey::trySetDwordValue(const std::wstring& valueName, const DWORD data) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_DWORD,
			reinterpret_cast<const BYTE*>(&data),
			sizeof(data)
		) };
	}

	inline RegResult RegKey::trySetQwordValue(const std::wstring& valueName, const ULONGLONG& data) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_QWORD,
			reinterpret_cast<const BYTE*>(&data),
			sizeof(data)
		) };
	}

	inline RegResult RegKey::trySetStringValue(const std::wstring& valueName, const std::wstring& data) {
		_ASSERTE(isValid());

		// String size including the terminating null character, in bytes
		const DWORD dataSize = Details::safeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_SZ,
			reinterpret_cast<const BYTE*>(data.c_str()),
			dataSize
		) };
	}
	
	inline RegResult RegKey::trySetExpandStringValue(const std::wstring& valueName, const std::wstring& data) {
		_ASSERTE(isValid());

		// String size including the terminating null character, in bytes
		const DWORD dataSize = Details::safeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_EXPAND_SZ,
			reinterpret_cast<const BYTE*>(data.c_str()),
			dataSize
		) };
	}

	inline RegResult RegKey::trySetMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data) {
		_ASSERTE(isValid());

		/**
		 * First, we have to build a double-null-terminated multi-string from the input data.
		 * 
		 * @note This is the reason why we can't mark this method as `noexcept`,
		 * since a dynamic allocation happens when creating the std::vector in `Details::buildMultiString`.
		 * And, if a dynamic allocation fails, an exception will be thrown.
		 */
		const std::vector<wchar_t> multiString = Details::buildMultiString(data);

		// Total size, in bytes, of the whole multi-string sequence
		const DWORD dataSize = Details::safeCastSizeToDword(multiString.size() * sizeof(wchar_t));

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_MULTI_SZ,
			reinterpret_cast<const BYTE*>(multiString.data()),
			dataSize	
		) };
	}

	inline RegResult RegKey::trySetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data) {
		_ASSERTE(isValid());

		// Total size, in bytes, of the binary data
		const DWORD dataSize = Details::safeCastSizeToDword(data.size());

		return RegResult{ ::RegSetValueExW(
			m_hKey,
			valueName.c_str(),
			0, // Reserved, must be zero
			REG_BINARY,
			data.data(),
			dataSize
		) };
	}

	inline DWORD RegKey::getDwordValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		DWORD data = 0; // To be read from the registry
		DWORD dataSize = sizeof(data); // Size of the data, in bytes

		constexpr DWORD kFlags = RRF_RT_REG_DWORD;
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey,
			nullptr, // No sub-key
			valueName.c_str(),
			kFlags,
			nullptr, // Type not required
			&data,
			&dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get DWORD value: RegGetValueW failed." };
		}

		return data;
	}

	inline ULONGLONG RegKey::getQwordValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		ULONGLONG data = 0; // To be read from the registry
		DWORD dataSize = sizeof(data); // Size of the data, in bytes

		constexpr DWORD kFlags = RRF_RT_REG_QWORD;
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey,
			nullptr, // No sub-key
			valueName.c_str(),
			kFlags,
			nullptr, // Type not required
			&data,
			&dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get QWORD value: RegGetValueW failed." };
		}

		return data;
	}

	inline std::wstring RegKey::getStringValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		std::wstring result; // To be read from the registry
		DWORD dataSize = 0; // Size of the string data, in bytes

		constexpr DWORD kFlags = RRF_RT_REG_SZ;

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Get the size of the result string
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot get the size of the string value: RegGetValueW failed." };
			}

			/**
			 * Allocate a string with the proper size (in bytes, including the null terminator)
			 * @note We have to convert the size from bytes to `wchar_t`s for `std::wstring::resize()`.
			 */
			result.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual string data
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				result.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get string value: RegGetValueW failed." };
		}

		// Remove the null terminator scribbled by `RegGetValueW` from the string
		result.resize((dataSize / sizeof(wchar_t)) - 1);

		return result;
	}

	inline std::wstring RegKey::getExpandStringValue(const std::wstring& valueName, const ExpandStringOption expandOption) const {
		_ASSERTE(isValid());

		std::wstring result; // To be read from the registry
		DWORD dataSize = 0; // Size of the expand string data, in bytes

		DWORD kFlags = RRF_RT_REG_EXPAND_SZ;

		// Adjust the flag for the `RegGetValueW` call, considering the expand option specified by the caller
		if (expandOption == ExpandStringOption::DontExpand) {
			kFlags |= RRF_NOEXPAND;
		}

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Get the size of the result string
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot get the size of the expand string value: RegGetValueW failed." };
			}

			/**
			 * Allocate a string of the proper size (in bytes, including the null terminator)
			 * @note We have to convert the size from bytes to `wchar_t`s for `std::wstring::resize()`.
			 */
			result.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual string data
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				result.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get expand string value: RegGetValueW failed." };
		}

		// Remove the null terminator scribbled by `RegGetValueW` from the `std::wstring`
		result.resize((dataSize / sizeof(wchar_t)) - 1);

		return result;
	}

	inline std::vector<std::wstring> RegKey::getMultiStringValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		// Room for the result multi-string, to be read from the registry
		std::vector<wchar_t> multiString;

		// Total size of the multi-string data, in bytes
		DWORD dataSize = 0;

		constexpr DWORD kFlags = RRF_RT_REG_MULTI_SZ;

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Request the size of the multi-string, in bytes
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot get the size of the multi-string value: RegGetValueW failed." };
			}

			/**
			 * Allocate room for the resulting multi-string
			 * @note that the `dataSize` is in bytes, but our `std::vector<wchar_t>::resize()` requires the size to be expressed in `wchar_t`s.
			 */
			multiString.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual multi-string data into the vector
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				multiString.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get the multi-string value: RegGetValueW failed." };
		}

		/**
		 * Resize the vector to the actual size returned by the last call to `RegGetValueW`.
		 * @note That the vector is a vector of `wchar_t`s, instead the size returned by `RegGetValueW` is in bytes,
		 * 		so we have to scale from bytes to `wchar_t` count.
		 */
		multiString.resize(dataSize / sizeof(wchar_t));

		// Convert the double-null-terminated multi-string to a vector of strings
		return Details::parseMultiString(multiString);
	}

	inline std::vector<BYTE> RegKey::getBinaryValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		// Room for the binary data, to be read from the registry
		std::vector<BYTE> binaryData;

		// Size of the binary data, in bytes
		DWORD dataSize = 0;

		constexpr DWORD kFlags = RRF_RT_REG_BINARY;

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA)
		{
			// Request the size of the binary data, in bytes
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot get the size of the binary value: RegGetValueW failed." };
			}

			// Allocate a buffer of the proper size to store the binary data
			binaryData.resize(dataSize);

			// Handle the special case of zero-length binary data:
			// If the binary data value in the registry is empty, just return an empty vector
			if (dataSize == 0) {
				_ASSERTE(binaryData.empty());
				return binaryData;
			}

			// Call `RegGetValueW` a second time to read the actual binary data into the vector
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				binaryData.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot get the binary data: RegGetValueW failed." };
		}

		// Resize the vector to the actual size returned by the last call to `RegGetValueW`
		binaryData.resize(dataSize);

		return binaryData;
	}

	inline RegExpected<DWORD> RegKey::tryGetDwordValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		DWORD data = 0; // To be read from the registry
		DWORD dataSize = sizeof(data); // Size of the data, in bytes

		constexpr DWORD kFlags = RRF_RT_REG_DWORD;
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey,
			nullptr, // No sub-key
			valueName.c_str(),
			kFlags,
			nullptr, // Type not required
			&data,
			&dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<DWORD>(errorCode);
		}

		return RegExpected<DWORD>{ data };
	}

	inline RegExpected<ULONGLONG> RegKey::tryGetQwordValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		ULONGLONG data = 0; // To be read from the registry
		DWORD dataSize = sizeof(data); // Size of the data, in bytes

		constexpr DWORD kFlags = RRF_RT_REG_QWORD;
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey,
			nullptr, // No sub-key
			valueName.c_str(),
			kFlags,
			nullptr, // Type not required
			&data,
			&dataSize
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<ULONGLONG>(errorCode);
		}

		return RegExpected<ULONGLONG>{ data };
	}

	inline RegExpected<std::wstring> RegKey::tryGetStringValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		constexpr DWORD kFlags = RRF_RT_REG_SZ;

		std::wstring result;
		DWORD dataSize;

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Get the size of the result string
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::wstring>(errorCode);
			}

			/**
			 * Allocate a string with the proper size (in bytes, including the null terminator)
			 * @note We have to convert the size from bytes to `wchar_t`s for `std::wstring::resize()`.
			 */
			result.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual string data
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				result.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::wstring>(errorCode);
		}

		// Remove the null terminator scribbled by `RegGetValueW` from the `std:wstring`
		result.resize((dataSize / sizeof(wchar_t)) - 1);

		return RegExpected<std::wstring>{ result };
	}

	inline RegExpected<std::wstring> RegKey::tryGetExpandStringValue(const std::wstring& valueName, const ExpandStringOption expandOption) const {
		_ASSERTE(isValid());

		DWORD kFlags = RRF_RT_REG_EXPAND_SZ;

		// Adjust the flag for the `RegGetValueW` call, considering the expand option specified by the caller
		if (expandOption == ExpandStringOption::DontExpand) {
			kFlags |= RRF_NOEXPAND;
		}

		std::wstring result;
		DWORD dataSize = 0; // Size of the expand string data, in bytes

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Get the size of the result string
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::wstring>(errorCode);
			}

			/**
			 * Allocate a string of the proper size (in bytes, including the null terminator)
			 * @note We have to convert the size from bytes to `wchar_t`s for `std::wstring::resize()`.
			 */
			result.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual string data
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				result.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::wstring>(errorCode);
		}

		// Remove the null terminator scribbled by `RegGetValueW` from the `std::wstring`
		result.resize((dataSize / sizeof(wchar_t)) - 1);

		return RegExpected<std::wstring>{ result };
	}

	inline RegExpected<std::vector<std::wstring>> RegKey::tryGetMultiStringValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		constexpr DWORD kFlags = RRF_RT_REG_MULTI_SZ;

		std::vector<wchar_t> data; // Room for the result multi-string, to be read from the registry
		DWORD dataSize = 0; // Size of the multi-string data, in bytes

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Request the size of the multi-string, in bytes
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::vector<std::wstring>>(errorCode);
			}

			/**
			 * Allocate room for the resulting multi-string
			 * @note that the `dataSize` is in bytes, but our `std::vector<wchar_t>::resize()` requires the size to be expressed in `wchar_t`s.
			 */
			data.resize(dataSize / sizeof(wchar_t));

			// Call `RegGetValueW` a second time to read the actual multi-string data into the vector
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				data.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::vector<std::wstring>>(errorCode);
		}

		/**
		 * Resize the vector to the actual size returned by the last call to `RegGetValueW`.
		 * @note That the vector is a vector of `wchar_t`s, instead the size returned by `RegGetValueW` is in bytes,
		 * 		so we have to scale from bytes to `wchar_t` count.
		 */
		data.resize(dataSize / sizeof(wchar_t));

		// Convert the double-null-terminated multi-string to a vector of strings
		return RegExpected<std::vector<std::wstring>>{ Details::parseMultiString(data) };
	}

	inline RegExpected<std::vector<BYTE>> RegKey::tryGetBinaryValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		constexpr DWORD kFlags = RRF_RT_REG_BINARY;

		std::vector<BYTE> data; // Room for the binary data, to be read from the registry
		DWORD dataSize = 0; // Size of the binary data, in bytes

		LSTATUS errorCode = ERROR_MORE_DATA;
		while (errorCode == ERROR_MORE_DATA) {
			// Request the size of the binary data, in bytes
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				nullptr, // Output buffer is not needed at this stage
				&dataSize
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::vector<BYTE>>(errorCode);
			}

			// Allocate a buffer of the proper size to store the binary data
			data.resize(dataSize);

			// Handle the special case of zero-length binary data:
			// If the binary data value in the registry is empty, just return an empty vector
			if (dataSize == 0) {
				_ASSERTE(data.empty());
				return RegExpected<std::vector<BYTE>>{ data };
			}

			// Call `RegGetValueW` a second time to read the actual binary data into the vector
			errorCode = ::RegGetValueW(
				m_hKey,
				nullptr, // No sub-key
				valueName.c_str(),
				kFlags,
				nullptr, // Type not required
				data.data(), // Output buffer
				&dataSize
			);
		}

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::vector<BYTE>>(errorCode);
		}

		// Resize the vector to the actual size returned by the last call to `RegGetValueW`
		data.resize(dataSize);

		return RegExpected<std::vector<BYTE>>{ data };
	}

	inline std::vector<std::wstring> RegKey::enumSubKeys() const {
		_ASSERTE(isValid());

		// Get some useful enumeration info, like the number of sub-keys & the maximum length of the sub-key names
		DWORD subKeyCount = 0;
		DWORD maxSubKeyNameLength = 0;

		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr, // No user-defined class
			nullptr, // No user-defined class size
			nullptr, // Reserved
			&subKeyCount,
			&maxSubKeyNameLength,
			nullptr, // No sub-key class length
			nullptr, // No value count
			nullptr, // No value name max length
			nullptr, // No max value length
			nullptr, // No security descriptor
			nullptr // No last write time
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "'RegQueryInfoKeyW' failed while preparing for sub-key enumeration." };
		}

		/**
		* @note According to the MSDN documentation, the size returned for the maximum length of sub-key names
		*		does not include the terminating null character. So, we have to add 1 to the size when allocating the buffer.
		*/
		maxSubKeyNameLength++;

		// Pre-allocate a buffer for the sub-key names
		auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLength);

		std::vector<std::wstring> subKeys; // To be filled with the sub-key names

		// Reserve enough room in the vector to accelerate the following insertion loop
		subKeys.reserve(subKeyCount);

		// Enumerate all the sub-keys
		for (DWORD i = 0; i < subKeyCount; i++) {
			// Get the name of the current sub-key
			DWORD subKeyNameLength = maxSubKeyNameLength;
			errorCode = ::RegEnumKeyExW(
				m_hKey,
				i,
				nameBuffer.get(),
				&subKeyNameLength,
				nullptr, // Reserved
				nullptr, // No class
				nullptr, // No class
				nullptr // No last write time
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot enumerate sub-keys: RegEnumKeyExW failed." };
			}

			/**
			* @note On success, `RegEnumKeyExW` writes the length of the sub-key name to `subKeyNameLength`
			* but doesn't account for the null terminator, so a `std::wstring` is built off that length.
			*/
			subKeys.emplace_back(nameBuffer.get(), subKeyNameLength);
		}

		return subKeys;
	}

	inline std::vector<std::pair<std::wstring, DWORD>> RegKey::enumValues() const {
		_ASSERTE(isValid());

		// Get some useful enumeration info, like the number of values & the maximum length of the value names
		DWORD valueCount = 0;
		DWORD maxValueNameLength = 0;

		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr, // No user-defined class
			nullptr, // No user-defined class size
			nullptr, // Reserved
			nullptr, // No sub-key count
			nullptr, // No sub-key max length
			nullptr, // No sub-key class length
			&valueCount,
			&maxValueNameLength,
			nullptr, // No max value length
			nullptr, // No security descriptor
			nullptr // No last write time
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "'RegQueryInfoKeyW' failed while preparing for value enumeration." };
		}

		/**
		* @note According to the MSDN documentation, the size returned for the maximum length of value names
		*		does not include the terminating null character. So, we have to add 1 to the size when allocating the buffer.
		*/
		maxValueNameLength++;

		// Pre-allocate a buffer for the value names
		auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLength);

		std::vector<std::pair<std::wstring, DWORD>> valueInfo; // To be filled with the value names & types

		// Reserve enough room in the vector to accelerate the following insertion loop
		valueInfo.reserve(valueCount);

		// Enumerate all the values
		for (DWORD i = 0; i < valueCount; i++) {
			// Get the name & type of the current value
			DWORD valueNameLength = maxValueNameLength;
			DWORD valueType = 0;
			errorCode = ::RegEnumValueW(
				m_hKey,
				i,
				nameBuffer.get(),
				&valueNameLength,
				nullptr, // Reserved
				&valueType,
				nullptr, // No data
				nullptr // No data size
			);

			if (errorCode != ERROR_SUCCESS) {
				throw RegException{ errorCode, "Cannot enumerate values: RegEnumValueW failed." };
			}

			/**
			* @note On success, `RegEnumValueW` writes the length of the value name to `valueNameLength`
			* but doesn't account for the null terminator, so a `std::wstring` is built off that length.
			*/
			valueInfo.emplace_back(std::wstring{ nameBuffer.get(), valueNameLength }, valueType);
		}

		return valueInfo;
	}

	inline bool RegKey::hasValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		// Invoke `RegGetValueW` to check if the value exists under the current key
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey, // The current key
			nullptr, // No sub-key - we're checking under the current key
			valueName.c_str(), // The name of the value to check
			RRF_RT_ANY, // No type restrictions on this value
			nullptr, // Type not required
			nullptr, // Output buffer is not needed
			nullptr // Data size is not needed
		);

		// If the function succeeds, the value exists
		if (errorCode == ERROR_SUCCESS) {
			return true;
		}
		else if (errorCode == ERROR_FILE_NOT_FOUND) {
			// If the function fails with this specific error code, the value does not exist
			return false;
		}
		else {
			// If the function fails with any other error code, throw an exception
			throw RegException{ errorCode, "Cannot check if the value exists: RegGetValueW failed." };
		}
	}

	inline bool RegKey::hasSubKey(const std::wstring& subKey) const {
		_ASSERTE(isValid());

		// Try to open the sub-key in read-only mode
		HKEY hSubKey = nullptr;
		LSTATUS errorCode = ::RegOpenKeyExW(
			m_hKey,
			subKey.c_str(),
			0, // Reserved
			KEY_READ,
			&hSubKey
		);

		// If we are able to open the sub-key, it does exist
		if (errorCode == ERROR_SUCCESS) {
			// Don't forget to close the sub-key
			::RegCloseKey(hSubKey);
			hSubKey = nullptr;

			return true;
		}
		else if ((errorCode == ERROR_FILE_NOT_FOUND) || (errorCode == ERROR_PATH_NOT_FOUND)) {
			// If the function fails with these specific error codes, the sub-key does not exist
			return false;
		}
		else {
			// If the function fails with any other error code, throw an exception
			throw RegException{ errorCode, "Cannot check if the sub-key exists: RegOpenKeyExW failed." };
		}
	}

	inline RegExpected<std::vector<std::wstring>> RegKey::tryEnumSubKeys() const {
		_ASSERTE(isValid());

		// Get some useful enumeration info, like the number of sub-keys & the maximum length of the sub-key names
		DWORD subKeyCount = 0;
		DWORD maxSubKeyNameLength = 0;

		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr, // No user-defined class
			nullptr, // No user-defined class size
			nullptr, // Reserved
			&subKeyCount,
			&maxSubKeyNameLength,
			nullptr, // No sub-key class length
			nullptr, // No value count
			nullptr, // No value name max length
			nullptr, // No max value length
			nullptr, // No security descriptor
			nullptr // No last write time
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::vector<std::wstring>>(errorCode);
		}

		/**
		* @note According to the MSDN documentation, the size returned for the maximum length of sub-key names
		*		does not include the terminating null character. So, we have to add 1 to the size when allocating the buffer.
		*/
		maxSubKeyNameLength++;

		// Pre-allocate a buffer for the sub-key names
		auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLength);

		std::vector<std::wstring> subKeysNames; // To be filled with the sub-key names

		// Reserve enough room in the vector to accelerate the following insertion loop
		subKeysNames.reserve(subKeyCount);

		// Enumerate all the sub-keys
		for (DWORD i = 0; i < subKeyCount; i++) {
			// Get the name of the current sub-key
			DWORD subKeyNameLength = maxSubKeyNameLength;
			errorCode = ::RegEnumKeyExW(
				m_hKey,
				i,
				nameBuffer.get(),
				&subKeyNameLength,
				nullptr, // Reserved
				nullptr, // No class
				nullptr, // No class
				nullptr // No last write time
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::vector<std::wstring>>(errorCode);
			}

			/**
			* @note On success, `RegEnumKeyExW` writes the length of the sub-key name to `subKeyNameLength`
			* but doesn't account for the null terminator, so a `std::wstring` is built off that length.
			*/
			subKeysNames.emplace_back(nameBuffer.get(), subKeyNameLength);
		}

		return RegExpected<std::vector<std::wstring>>{ subKeysNames };
	}

	inline RegExpected<std::vector<std::pair<std::wstring, DWORD>>> RegKey::tryEnumValues() const {
		_ASSERTE(isValid());

		// Get some useful enumeration info, like the number of values & the maximum length of the value names
		DWORD valueCount = 0;
		DWORD maxValueNameLength = 0;

		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr, // No user-defined class
			nullptr, // No user-defined class size
			nullptr, // Reserved
			nullptr, // No sub-key count
			nullptr, // No sub-key max length
			nullptr, // No sub-key class length
			&valueCount,
			&maxValueNameLength,
			nullptr, // No max value length
			nullptr, // No security descriptor
			nullptr // No last write time
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<std::vector<std::pair<std::wstring, DWORD>>>(errorCode);
		}

		/**
		* @note According to the MSDN documentation, the size returned for the maximum length of value names
		*		does not include the terminating null character. So, we have to add 1 to the size when allocating the buffer.
		*/
		maxValueNameLength++;

		// Pre-allocate a buffer for the value names
		auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLength);

		std::vector<std::pair<std::wstring, DWORD>> valueInfo; // To be filled with the value names & types
		
		// Reserve enough room in the vector to accelerate the following insertion loop
		valueInfo.reserve(valueCount);

		// Enumerate all the values
		for (DWORD i = 0; i < valueCount; i++) {
			// Get the name & type of the current value
			DWORD valueNameLength = maxValueNameLength;
			DWORD valueType = 0;
			errorCode = ::RegEnumValueW(
				m_hKey,
				i,
				nameBuffer.get(),
				&valueNameLength,
				nullptr, // Reserved
				&valueType,
				nullptr, // No data
				nullptr // No data size
			);

			if (errorCode != ERROR_SUCCESS) {
				return Details::makeRegExpectedWithError<std::vector<std::pair<std::wstring, DWORD>>>(errorCode);
			}

			/**
			* @note On success, `RegEnumValueW` writes the length of the value name to `valueNameLength`
			* but doesn't account for the null terminator, so a `std::wstring` is built off that length.
			*/
			valueInfo.emplace_back(std::wstring{ nameBuffer.get(), valueNameLength }, valueType);
		}

		return RegExpected<std::vector<std::pair<std::wstring, DWORD>>>{ valueInfo };
	}

	inline RegExpected<bool> RegKey::tryHasValue(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		// Invoke `RegGetValueW` to check if the value exists under the current key
		LSTATUS errorCode = ::RegGetValueW(
			m_hKey, // The current key
			nullptr, // No sub-key - we're checking under the current key
			valueName.c_str(), // The name of the value to check
			RRF_RT_ANY, // No type restrictions on this value
			nullptr, // Type not required
			nullptr, // Output buffer is not needed
			nullptr // Data size is not needed
		);

		// If the function succeeds, the value exists
		if (errorCode == ERROR_SUCCESS) {
			return RegExpected<bool>{ true };
		}
		else if (errorCode == ERROR_FILE_NOT_FOUND) {
			// If the function fails with this specific error code, the value does not exist
			return RegExpected<bool>{ false };
		}
		else {
			// If the function fails with any other error code, return an error
			return Details::makeRegExpectedWithError<bool>(errorCode);
		}
	}

	inline RegExpected<bool> RegKey::tryHasSubKey(const std::wstring& subKey) const {
		_ASSERTE(isValid());

		// Try to open the sub-key in read-only mode
		HKEY hSubKey = nullptr;
		LSTATUS errorCode = ::RegOpenKeyExW(
			m_hKey,
			subKey.c_str(),
			0, // Reserved
			KEY_READ,
			&hSubKey
		);

		// If we are able to open the sub-key, it does exist
		if (errorCode == ERROR_SUCCESS) {
			// Don't forget to close the sub-key
			::RegCloseKey(hSubKey);
			hSubKey = nullptr;

			return RegExpected<bool>{ true };
		}
		else if ((errorCode == ERROR_FILE_NOT_FOUND) || (errorCode == ERROR_PATH_NOT_FOUND)) {
			// If the function fails with these specific error codes, the sub-key does not exist
			return RegExpected<bool>{ false };
		}
		else {
			// If the function fails with any other error code, return an error
			return Details::makeRegExpectedWithError<bool>(errorCode);
		}
	}

	inline DWORD RegKey::queryValueType(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		DWORD typeId = 0; // Will be returned by `RegQueryValueExW`
		LSTATUS errorCode = ::RegQueryValueExW(
			m_hKey,
			valueName.c_str(),
			nullptr, // Reserved
			&typeId,
			nullptr, // Data not required
			nullptr // Data size not required
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot query the value type: RegQueryValueExW failed." };
		}

		return typeId;
	}

	inline RegExpected<DWORD> RegKey::tryQueryValueType(const std::wstring& valueName) const {
		_ASSERTE(isValid());

		DWORD typeId = 0; // Will be returned by `RegQueryValueExW`
		LSTATUS errorCode = ::RegQueryValueExW(
			m_hKey,
			valueName.c_str(),
			nullptr, // Reserved
			&typeId,
			nullptr, // Data not required
			nullptr // Data size not required
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<DWORD>(errorCode);
		}

		return RegExpected<DWORD>{ typeId };
	}

	inline RegKey::InfoKey RegKey::queryInfoKey() const {
		_ASSERTE(isValid());

		InfoKey infoKey{};
		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr,
			nullptr,
			nullptr,
			&(infoKey.NumberOfSubKeys),
			nullptr,
			nullptr,
			&(infoKey.NumberOfValues),
			nullptr,
			nullptr,
			nullptr,
			&(infoKey.LastWriteTime)
		);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot query key information: RegQueryInfoKeyW failed." };
		}

		return infoKey;
	}

	inline RegExpected<RegKey::InfoKey> RegKey::tryQueryInfoKey() const {
		_ASSERTE(isValid());

		InfoKey infoKey{};
		LSTATUS errorCode = ::RegQueryInfoKeyW(
			m_hKey,
			nullptr,
			nullptr,
			nullptr,
			&(infoKey.NumberOfSubKeys),
			nullptr,
			nullptr,
			&(infoKey.NumberOfValues),
			nullptr,
			nullptr,
			nullptr,
			&(infoKey.LastWriteTime)
		);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<InfoKey>(errorCode);
		}

		return RegExpected<InfoKey>{ infoKey };
	}

	inline RegKey::KeyReflection RegKey::queryReflectionKey() const {
		BOOL isReflectionDisabled = FALSE;
		LSTATUS errorCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot query the reflection status of the key: RegQueryReflectionKey failed." };
		}

		return isReflectionDisabled ? KeyReflection::ReflectionDisabled : KeyReflection::ReflectionEnabled;
	}

	inline RegExpected<RegKey::KeyReflection> RegKey::tryQueryReflectionKey() const {
		BOOL isReflectionDisabled = FALSE;
		LSTATUS errorCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);

		if (errorCode != ERROR_SUCCESS) {
			return Details::makeRegExpectedWithError<KeyReflection>(errorCode);
		}

		return RegExpected<KeyReflection>{ isReflectionDisabled ? KeyReflection::ReflectionDisabled : KeyReflection::ReflectionEnabled };
	}

	inline void RegKey::deleteValue(const std::wstring& valueName) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegDeleteKeyW(m_hKey, valueName.c_str());

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot delete the value: RegDeleteValueW failed." };
		}
	}

	inline RegResult RegKey::tryDeleteValue(const std::wstring& valueName) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegDeleteValueW(m_hKey, valueName.c_str()) };
	}

	inline void RegKey::deleteKey(const std::wstring& subKey, const REGSAM desiredAccess) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot delete the sub-key: RegDeleteKeyExW failed." };
		}
	}

	inline RegResult RegKey::tryDeleteKey(const std::wstring& subKey, const REGSAM desiredAccess) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0) };
	}

	inline void RegKey::deleteTree(const std::wstring& subKey) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegDeleteTreeW(m_hKey, subKey.c_str());

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot delete the sub-key tree: RegDeleteTreeW failed." };
		}
	}

	inline RegResult RegKey::tryDeleteTree(const std::wstring& subKey) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegDeleteTreeW(m_hKey, subKey.c_str()) };
	}

	inline void RegKey::copyTree(const std::wstring& srcSubKey, const RegKey& dstKey) {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegCopyTreeW(m_hKey, srcSubKey.c_str(), dstKey.get());

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot copy the sub-key tree: RegCopyTreeW failed." };
		}
	}

	inline RegResult RegKey::tryCopyTree(const std::wstring& srcSubKey, const RegKey& dstKey) noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegCopyTreeW(m_hKey, srcSubKey.c_str(), dstKey.get()) };
	}

	inline void RegKey::flushKey() {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegFlushKey(m_hKey);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot flush the key: RegFlushKey failed." };
		}
	}

	inline RegResult RegKey::tryFlushKey() noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegFlushKey(m_hKey) };
	}

	inline void RegKey::loadKey(const std::wstring& subKey, const std::wstring& filename) {
		close(); // Close the current key before loading a new one

		LSTATUS errorCode = ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str());

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot load the key: RegLoadKeyW failed." };
		}
	}

	inline RegResult RegKey::tryLoadKey(const std::wstring& subKey, const std::wstring& filename) noexcept {
		close(); // Close the current key before loading a new one

		return RegResult{ ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str()) };
	}

	inline void RegKey::saveKey(const std::wstring& filename, SECURITY_ATTRIBUTES* const securityAttributes) const {
		_ASSERTE(isValid());

		LSTATUS errorCode = ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot save the key: RegSaveKeyW failed." };
		}
	}

	inline RegResult RegKey::trySaveKey(const std::wstring& filename, SECURITY_ATTRIBUTES* const securityAttributes) const noexcept {
		_ASSERTE(isValid());

		return RegResult{ ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes) };
	}

	inline void RegKey::enableReflectionKey() {
		LSTATUS errorCode = ::RegEnableReflectionKey(m_hKey);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot enable reflection for the key: RegEnableReflectionKey failed." };
		}
	}

	inline RegResult RegKey::tryEnableReflectionKey() noexcept {
		return RegResult{ ::RegEnableReflectionKey(m_hKey) };
	}

	inline void RegKey::disableReflectionKey() {
		LSTATUS errorCode = ::RegDisableReflectionKey(m_hKey);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot disable reflection for the key: RegDisableReflectionKey failed." };
		}
	}

	inline RegResult RegKey::tryDisableReflectionKey() noexcept {
		return RegResult{ ::RegDisableReflectionKey(m_hKey) };
	}

	inline void RegKey::connectRegistry(const std::wstring& machineName, const HKEY hKeyPredefined) {
		close(); // Safely close any previously opened key

		HKEY hKeyResult = nullptr;

		LSTATUS errorCode = ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult);

		if (errorCode != ERROR_SUCCESS) {
			throw RegException{ errorCode, "Cannot connect to the registry: RegConnectRegistryW failed." };
		}

		// Take ownership of the newly opened key
		m_hKey = hKeyResult;
	}

	inline RegResult RegKey::tryConnectRegistry(const std::wstring& machineName, const HKEY hKeyPredefined) noexcept {
		close(); // Safely close any previously opened key

		HKEY hKeyResult = nullptr;

		RegResult errorCode{ ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult) };

		if (errorCode.failed()) {
			return errorCode;
		}

		// Take ownership of the newly opened key
		m_hKey = hKeyResult;

		_ASSERTE(errorCode.isOk());
		return errorCode;
	}

	inline std::wstring RegKey::regTypeToString(const DWORD regType) {
		switch(regType) {
			case REG_SZ: return L"REG_SZ";
			case REG_EXPAND_SZ: return L"REG_EXPAND_SZ";
			case REG_MULTI_SZ: return L"REG_MULTI_SZ";
			case REG_DWORD: return L"REG_DWORD";
			case REG_QWORD: return L"REG_QWORD";
			case REG_BINARY: return L"REG_BINARY";
			default: return L"Unknown";
		}
	}

	// RegException inline methods
	inline RegException::RegException(const LSTATUS errorCode, const char* const message) : std::system_error{ errorCode, std::system_category(), message } {}
	inline RegException::RegException(const LSTATUS errorCode, const std::string& message) : std::system_error{ errorCode, std::system_category(), message } {}

	// RegResult inline methods
	inline RegResult::RegResult(const LSTATUS result) noexcept : m_result{ result } {}

	inline bool RegResult::isOk() const noexcept {
		return m_result == ERROR_SUCCESS;
	}

	inline bool RegResult::failed() const noexcept {
		return m_result != ERROR_SUCCESS;
	}

	inline RegResult::operator bool() const noexcept {
		return isOk();
	}

	inline LSTATUS RegResult::code() const noexcept {
		return m_result;
	}

	inline std::wstring RegResult::message() const {
		return message(MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT));
	}

	inline std::wstring RegResult::message(const DWORD languageId) const {
		// Invoke `FormatMessageW` to get the error message from windows
		Details::ScopedLocalFree<wchar_t> messagePtr;

		DWORD errorCode = ::FormatMessageW(
			FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			nullptr, // Message is not associated with any module
			m_result,
			languageId,
			reinterpret_cast<LPWSTR>(messagePtr.addressOf()),
			0, // Minimum buffer size
			nullptr // No arguments
		);

		if (errorCode == 0) {
			// Return an empty string if `FormatMessageW` fails
			return std::wstring{};
		}

		// Safely copy the c-string returned by `FormatMessageW` into a `std::wstring`
		return std::wstring{ messagePtr.get() };
	}

	// RegExpected inline methods
	template<typename T>
	inline RegExpected<T>::RegExpected(const RegResult& errorCode) noexcept : m_var{ errorCode } {}

	template<typename T>
	inline RegExpected<T>::RegExpected(const T& value) : m_var{ value } {}

	template<typename T>
	inline RegExpected<T>::RegExpected(T&& value) : m_var{ std::move(value) } {}

	template<typename T>
	inline RegExpected<T>::operator bool() const noexcept {
		return isValid();
	}

	template<typename T>
	inline bool RegExpected<T>::isValid() const noexcept {
		return std::holds_alternative<T>(m_var);
	}

	template<typename T>
	inline const T& RegExpected<T>::getValue() const {
		// Check that the object stores a valid value
		_ASSERTE(isValid());

		// If the object is in a valid state, the variant stores an instance of `T`
		return std::get<T>(m_var);
	}

	template<typename T>
	inline RegResult RegExpected<T>::getError() const {
		// Check that the object stores an error code
		_ASSERTE(!isValid());

		// If the object is in an invalid state, the variant stores a `RegResult`
		// which represents an error code from the Windows Registry API
		return std::get<RegResult>(m_var);
	}
} // namespace WinReg