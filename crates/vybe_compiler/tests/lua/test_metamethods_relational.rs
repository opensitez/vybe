//! Relational metamethods: __lt, __le, __eq semantics (Lua 5.x §2.4)

lua_print! {
    relational_lt => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) < W(10))\n",
        "true"
    },
    relational_le => {
        "local mt = {__le = function(a, b) return a.v <= b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) <= W(5))\n",
        "true"
    },
    relational_eq => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) == W(5))\n",
        "true"
    },
    relational_eq_diff => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) == W(6))\n",
        "false"
    },
    relational_gt => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(10) > W(5))\n",
        "true"
    },
    relational_ge => {
        "local mt = {__le = function(a, b) return a.v <= b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(10) >= W(5))\n",
        "true"
    },
}
