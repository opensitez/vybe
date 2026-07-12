//! Memoization tables and dynamic caches utilizing weak tables (Lua 5.x §2.5.4)

lua_print! {
    memoize_weak_kv => {
        "local cache = setmetatable({}, {__mode=\"kv\"})\nlocal function get_cached(k)\n  if not cache[k] then cache[k] = {val = k.val * 2} end\n  return cache[k]\nend\nlocal key = {val = 10}\nlocal v1 = get_cached(key)\nlocal v2 = get_cached(key)\nprint(v1.val == v2.val)\n",
        "true"
    },
    weak_keys_gc => {
        "local t = setmetatable({}, {__mode=\"k\"})\nlocal key = {}\nt[key] = \"val\"\nprint(t[key])\n",
        "val"
    },
    weak_values_gc => {
        "local t = setmetatable({}, {__mode=\"v\"})\nlocal key = {}\nlocal val = {}\nt[key] = val\nprint(t[key] == val)\n",
        "true"
    },
}
