//! `__index` as function vs table, lookup chains, and `__newindex` proxies (Lua 5.x §2.4)

lua_print! {
    index_func => {
        "local t = setmetatable({}, {\n  __index = function(_, k) return k:upper() end\n})\nprint(t.hello)\n",
        "HELLO"
    },
    index_chain_two => {
        "local base = {x = 10}\nlocal mid = setmetatable({}, {__index = base})\nlocal top = setmetatable({}, {__index = mid})\nprint(top.x)\n",
        "10"
    },
    index_chain_three => {
        "local L1 = {v = 42}\nlocal L2 = setmetatable({}, {__index = L1})\nlocal L3 = setmetatable({}, {__index = L2})\nlocal L4 = setmetatable({}, {__index = L3})\nprint(L4.v)\n",
        "42"
    },
    newindex_proxy => {
        "local store = {}\nlocal proxy = setmetatable({}, {\n  __newindex = function(_, k, v) store[k] = v * 2 end,\n  __index = store\n})\nproxy.x = 5\nprint(proxy.x)\n",
        "10"
    },
    index_func_args => {
        "local log = nil\nlocal t = setmetatable({}, {\n  __index = function(tbl, k)\n    log = type(tbl) .. \",\" .. k\n    return 0\n  end\n})\n_ = t.foo\nprint(log)\n",
        "table,foo"
    },
    newindex_missing_only => {
        "local called = 0\nlocal t = setmetatable({x = 1}, {\n  __newindex = function(tbl, k, v)\n    called = called + 1\n    rawset(tbl, k, v)\n  end\n})\nt.x = 99   -- existing, no __newindex\nt.y = 10   -- new, triggers\nprint(called, t.x, t.y)\n",
        "1 99 10"
    },
    index_table_proto => {
        "local proto = {kind = \"base\", greet = function(self) return \"hi \" .. self.name end}\nlocal obj = setmetatable({name = \"world\"}, {__index = proto})\nprint(obj:greet())\n",
        "hi world"
    },
    index_exist_shortcircuit => {
        "local called = false\nlocal t = setmetatable({k = 99}, {\n  __index = function() called = true; return 0 end\n})\nlocal v = t.k\nprint(v, called)\n",
        "99 false"
    },
}
