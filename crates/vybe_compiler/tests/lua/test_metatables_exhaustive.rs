//! Metatables exhaustive tests: arithmetic, bitwise, relational overrides and get/set boundary checks (Lua 5.x §2.4)

lua_print! {
    meta_exh_add => {
        "local mt = {__add = function(a, b) return a.v + b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) + W(10))\n",
        "15"
    },
    meta_exh_sub => {
        "local mt = {__sub = function(a, b) return a.v - b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(15) - W(5))\n",
        "10"
    },
    meta_exh_mul => {
        "local mt = {__mul = function(a, b) return a.v * b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) * W(10))\n",
        "50"
    },
    meta_exh_div => {
        "local mt = {__div = function(a, b) return a.v / b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(10) / W(4))\n",
        "2.5"
    },
    meta_exh_mod => {
        "local mt = {__mod = function(a, b) return a.v % b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(10) % W(3))\n",
        "1"
    },
    meta_exh_pow => {
        "local mt = {__pow = function(a, b) return a.v ^ b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(2) ^ W(3))\n",
        "8.0"
    },
    meta_exh_unm => {
        "local mt = {__unm = function(a) return -a.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(-W(5))\n",
        "-5"
    },
    meta_exh_idiv => {
        "local mt = {__idiv = function(a, b) return a.v // b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(10) // W(3))\n",
        "3"
    },
    meta_exh_band => {
        "local mt = {__band = function(a, b) return a.v & b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(6) & W(3))\n",
        "2"
    },
    meta_exh_bor => {
        "local mt = {__bor = function(a, b) return a.v | b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(6) | W(3))\n",
        "7"
    },
    meta_exh_bxor => {
        "local mt = {__bxor = function(a, b) return a.v ~ b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(6) ~ W(3))\n",
        "5"
    },
    meta_exh_bnot => {
        "local mt = {__bnot = function(a) return ~a.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(~W(6))\n",
        "-7"
    },
    meta_exh_shl => {
        "local mt = {__shl = function(a, b) return a.v << b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(1) << W(3))\n",
        "8"
    },
    meta_exh_shr => {
        "local mt = {__shr = function(a, b) return a.v >> b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(8) >> W(3))\n",
        "1"
    },
    meta_exh_eq => {
        "local mt = {__eq = function(a, b) return a.v == b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) == W(5))\n",
        "true"
    },
    meta_exh_lt => {
        "local mt = {__lt = function(a, b) return a.v < b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) < W(10))\n",
        "true"
    },
    meta_exh_le => {
        "local mt = {__le = function(a, b) return a.v <= b.v end}\nlocal function W(v) return setmetatable({v=v}, mt) end\nprint(W(5) <= W(5))\n",
        "true"
    },
}
