//! `__metatable` field — protecting metatables from access/change (Lua 5.x §6.1)

lua_print! {
    metatable_protected => {
        "local mt = {__metatable = \"protected\"}\nlocal t = setmetatable({}, mt)\nprint(getmetatable(t))\n",
        "protected"
    },
    metatable_protected_set_raises => {
        "local mt = {__metatable = \"lock\"}\nlocal t = setmetatable({}, mt)\nlocal ok = pcall(setmetatable, t, {})\nprint(ok)\n",
        "false"
    },
    metatable_unprotected => {
        "local mt = {}\nlocal t = setmetatable({}, mt)\nprint(type(getmetatable(t)))\n",
        "table"
    },
    metatable_protected_val_table => {
        "local guard = {}\nlocal mt = {__metatable = guard}\nlocal t = setmetatable({}, mt)\nprint(getmetatable(t) == guard)\n",
        "true"
    },
    metatable_nil => {
        "print(tostring(getmetatable(nil)))\n",
        "nil"
    },
    metatable_number => {
        "print(tostring(getmetatable(42)))\n",
        "nil"
    },
    metatable_string => {
        "print(type(getmetatable(\"\")))\n",
        "table"
    },
    setmetatable_ret => {
        "local t = {}\nlocal r = setmetatable(t, {})\nprint(r == t)\n",
        "true"
    },
    setmetatable_nil => {
        "local t = setmetatable({}, {})\nsetmetatable(t, nil)\nprint(tostring(getmetatable(t)))\n",
        "nil"
    },
    metatable_shared => {
        "local mt = {__index = function() return 99 end}\nlocal a = setmetatable({}, mt)\nlocal b = setmetatable({}, mt)\nmt.__index = function() return 42 end\nprint(a.x, b.x)\n",
        "42\t42"
    },
}
