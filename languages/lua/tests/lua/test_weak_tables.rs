//! Weak tables — `__mode` (Lua 5.x manual §2.5.2).

lua_print! {
    weak_keys_mode_recorded_in_metatable => {
        "local m = {__mode = \"k\"}\nlocal t = setmetatable({}, m)\nprint(getmetatable(t).__mode)\n",
        "k"
    },
    weak_values_mode_recorded_in_metatable => {
        "local t = setmetatable({}, {__mode = \"v\"})\nprint(getmetatable(t).__mode)\n",
        "v"
    },
    weak_both_mode_kv => {
        "local t = setmetatable({}, {__mode = \"kv\"})\nprint(getmetatable(t).__mode)\n",
        "kv"
    },
    weak_table_still_allows_read_write => {
        "local t = setmetatable({}, {__mode = \"k\"})\nt.x = 1\nprint(t.x)\n",
        "1"
    },
}
