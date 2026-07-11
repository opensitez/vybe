//! Custom searchpath searching, package loaders, searchers list (Lua 5.2+ manual §6.3)

lua_print! {
    searchpath_finds => {
        "local path = package.searchpath and type(package.searchpath(\"package\", package.path))\nprint(path == \"string\" or path == \"nil\")\n",
        "true"
    },
    searchpath_missing => {
        "local p = package.searchpath and package.searchpath(\"missing_module\", package.path)\nprint(tostring(p))\n",
        "nil"
    },
    searchers_type => {
        "local searchers = package.searchers or package.loaders\nprint(type(searchers))\n",
        "table"
    },
    searchers_count => {
        "local searchers = package.searchers or package.loaders\nprint(#searchers >= 1)\n",
        "true"
    },
}
