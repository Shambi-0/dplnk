cmake_minimum_required(VERSION 3.25)

project(
	dplnk
	VERSION 0.0.1
	DESCRIPTION "A cross-platform deep linking library for c++"
	LANGUAGES CXX
)

set(CMAKE_DISABLE_IN_SOURCE_BUILD ON)

if (POLICY CMP0141)
  cmake_policy(SET CMP0141 NEW)
  set(CMAKE_MSVC_DEBUG_INFORMATION_FORMAT "$<IF:$<AND:$<C_COMPILER_ID:MSVC>,$<CXX_COMPILER_ID:MSVC>>,$<$<CONFIG:Debug,RelWithDebInfo>:EditAndContinue>,$<$<CONFIG:Debug,RelWithDebInfo>:ProgramDatabase>>")
endif()

file(GLOB SOURCES "src/*.cpp" "src/*.hpp")
file(GLOB HEADERS "include/${PROJECT_NAME}/*.h" "include/${PROJECT_NAME}/*.hpp")

add_library(${PROJECT_NAME} ${SOURCES} ${HEADERS})

set_target_properties(${PROJECT_NAME} PROPERTIES
	INTERPROCEDURAL_OPTIMIZATION_RELEASE TRUE
	INTERPROCEDURAL_OPTIMIZATION_PROFILE TRUE
)

target_compile_definitions(${PROJECT_NAME} PUBLIC RMLUI_USE_CUSTOM_RTTI=1)
set_property(TARGET ${PROJECT_NAME} PROPERTY CXX_STANDARD 23)

target_include_directories(${PROJECT_NAME}
PUBLIC
	"include"

PRIVATE
	"include/${PROJECT_NAME}/"
)

include(GNUInstallDirs)

install(TARGETS ${PROJECT_NAME}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    PUBLIC_HEADER DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)