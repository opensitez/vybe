//! `package` library — module searchers (Lua 5.2+ manual §6.3).

lua_print! {
    package_path_is_string => {
        "print(type(package.path) == \"string\")\n",
        "true"
    },
    package_cpath_is_string => {
        "print(type(package.cpath) == \"string\")\n",
        "true"
    },
    package_loaded_is_table => {
        "print(type(package.loaded) == \"table\")\n",
        "true"
    },
    package_searchers_is_table => {
        "print(type(package.searchers) == \"table\" or type(package.loaders) == \"table\")\n",
        "true"
    },
    package_preload_is_table => {
        "print(type(package.preload) == \"table\")\n",
        "true"
    },
    require_caches_module_in_loaded => {
        "package.loaded.fake = {v = 1}\nlocal m = require(\"fake\")\nprint(m.v)\n",
        "1"
    },
    package_searchpath_returns_filename_or_nil => {
        "local p = package.searchpath and package.searchpath(\"init\", package.path)\nprint(p == nil or type(p) == \"string\")\n",
        "true"
    },
    loaded_table_reuses_same_table => {
        "package.loaded.mod = {x = 1}\nlocal a = require(\"mod\")\na.x = 2\nlocal b = require(\"mod\")\nprint(b.x)\n",
        "2"
    },
}
