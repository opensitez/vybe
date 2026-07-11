//! Metatables extended tests — index, newindex, call, tostring, metatable guard, len (Lua 5.x §2.4)

lua_print! {
    meta_index_table => {
        "local parent = {x=10}\nlocal child = setmetatable({}, {__index = parent})\nprint(child.x)\n",
        "10"
    },
    meta_index_function => {
        "local t = setmetatable({}, {__index = function(tbl, key) return key .. \"!\" end})\nprint(t.hello)\n",
        "hello!"
    },
    meta_newindex_table => {
        "local parent = {}\nlocal child = setmetatable({}, {__newindex = parent})\nchild.x = 42\nprint(parent.x, tostring(rawget(child, \"x\")))\n",
        "42\tnil"
    },
    meta_newindex_function => {
        "local log = nil\nlocal t = setmetatable({}, {__newindex = function(tbl, k, v) log = k .. \"=\" .. v end})\nt.score = 100\nprint(log)\n",
        "score=100"
    },
    meta_call_functor => {
        "local mt = {__call = function(self, a, b) return a + b end}\nlocal obj = setmetatable({}, mt)\nprint(obj(5, 10))\n",
        "15"
    },
    meta_tostring_custom => {
        "local mt = {__tostring = function(self) return \"custom_str\" end}\nlocal obj = setmetatable({}, mt)\nprint(tostring(obj))\n",
        "custom_str"
    },
    meta_len_table => {
        "local mt = {__len = function() return 99 end}\nlocal t = setmetatable({}, mt)\nprint(#t)\n",
        "99"
    },
    meta_len_string => {
        "print(#\"abc\")\n",
        "3"
    },
    meta_metatable_guard => {
        "local mt = {__metatable = \"locked\"}\nlocal t = setmetatable({}, mt)\nprint(getmetatable(t))\n",
        "locked"
    },
    meta_setmetatable_locked_fails => {
        "local mt = {__metatable = \"locked\"}\nlocal t = setmetatable({}, mt)\nlocal ok = pcall(setmetatable, t, {})\nprint(ok)\n",
        "false"
    },
    meta_setmetatable_nil => {
        "local t = setmetatable({}, {__index={x=1}})\nsetmetatable(t, nil)\nprint(tostring(t.x))\n",
        "nil"
    },
}
