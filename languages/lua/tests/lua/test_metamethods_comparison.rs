//! Metatables: `__lt`, `__le`, `__eq` comparison metamethods (Lua 5.x §2.4, §6.1)

lua_print! {
    lt_metamethod => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(1) < W(2))\n",
        "true"
    },
    lt_metamethod_false => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(5) < W(2))\n",
        "false"
    },
    le_metamethod => {
        "local mt = {__le = function(a, b) return a.v <= b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(3) <= W(3))\n",
        "true"
    },
    le_via_lt_fallback => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(2) <= W(3))\n",
        "true"
    },
    eq_metamethod => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(5) == W(5))\n",
        "true"
    },
    eq_metamethod_false => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(5) == W(6))\n",
        "false"
    },
    gt_uses_lt => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(5) > W(3))\n",
        "true"
    },
    ge_uses_le => {
        "local mt = {__le = function(a, b) return a.v <= b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(5) >= W(5))\n",
        "true"
    },
    ne_uses_eq => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(1) ~= W(2))\n",
        "true"
    },
    sort_uses_lt => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nlocal t = {W(3), W(1), W(2)}\ntable.sort(t)\nprint(t[1].v .. \",\" .. t[2].v .. \",\" .. t[3].v)\n",
        "1,2,3"
    } }
