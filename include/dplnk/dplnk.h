#pragma once

#include <optional>
#include <string>
#include <map>

namespace dplnk {
	struct options {
		std::string protocol;
		std::optional<std::map<std::string, std::string>> d;
	};
	
	void dplnk(const std::string& path, options options);
} // namespace dplnk