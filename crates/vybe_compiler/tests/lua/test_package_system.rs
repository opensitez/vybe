//! package table: loaded cache, preload, path searching (Lua 5.x §6.3)

lua_print! {
    sys_package_loaded_tbl => {
        "print(type(package.loaded))\n",
        "table"
    },
    sys_package_preload_tbl => {
        "print(type(package.preload))\n",
        "table"
    },
    sys_package_path_str => {
        "print(type(package.path))\n",
        "string"
    },
    require_preload_mod => {
        "package.preload[\"mymod\"] = function() return {val=42} end\nlocal m = require(\"mymod\")\nprint(m.val)\n",
        "42"
    },
    require_cache_mod => {
        "package.preload[\"cached\"] = function() return {n=0} end\nlocal a = require(\"cached\")\na.n = 99\nlocal b = require(\"cached\")\nprint(b.n)\n",
        "99"
    },
    package_loaded_directly => {
        "package.loaded[\"fake\"] = {x=7}\nlocal m = require(\"fake\")\nprint(m.x)\n",
        "7"
    },
    require_bool_preload => {
        "package.preload[\"bool_mod\"] = function() return true end\nprint(require(\"bool_mod\"))\n",
        "true"
    },
    sys_package_searchers_tbl => {
        "print(type(package.searchers or package.loaders))\n",
        "table"
    },
}
