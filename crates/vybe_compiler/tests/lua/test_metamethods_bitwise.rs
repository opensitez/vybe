//! Bitwise metamethods — `__band`, `__bor`, `__bxor`, `__bnot`, `__shl`, `__shr` (Lua 5.3+)

lua_print! {
    band_metamethod => {
        "local mt = {__band = function(a, b) return setmetatable({v = a.v & b.v}, getmetatable(a)) end}\nmt.__index = mt\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint(W(0xFF).__band(W(0xFF), W(0x0F)).v)\n",
        "15"
    },
    bor_metamethod => {
        "local mt = {__bor = function(a, b) return {v = a.v | b.v} end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nlocal r = W(0x01) | W(0x02)\nprint(r.v)\n",
        "3"
    },
    bxor_metamethod => {
        "local mt = {__bxor = function(a, b) return {v = a.v ~ b.v} end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nlocal r = W(0xFF) ~ W(0x0F)\nprint(r.v)\n",
        "240"
    },
    bnot_metamethod => {
        "local mt = {__bnot = function(a) return {v = ~a.v & 0xFF} end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint((~W(0)).v)\n",
        "255"
    },
    shl_metamethod => {
        "local mt = {__shl = function(a, n) return {v = a.v << n} end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint((W(1) << 4).v)\n",
        "16"
    },
    shr_metamethod => {
        "local mt = {__shr = function(a, n) return {v = a.v >> n} end}\nlocal function W(n) return setmetatable({v=n}, mt) end\nprint((W(256) >> 4).v)\n",
        "16"
    },
    native_band => {
        "print(0xFF & 0x0F)\n",
        "15"
    },
    native_bor => {
        "print(0xF0 | 0x0F)\n",
        "255"
    },
    native_bxor => {
        "print(0xFF ~ 0x0F)\n",
        "240"
    },
    native_bnot => {
        "print(~0)\n",
        "-1"
    },
    native_shl => {
        "print(1 << 1)\n",
        "2"
    },
    native_shr => {
        "print(8 >> 1)\n",
        "4"
    },
    bitwise_shift_zero => {
        "print(42 << 0)\n",
        "42"
    },
    bitwise_and_zero => {
        "print(0xFFFF & 0)\n",
        "0"
    },
    bitwise_or_zero => {
        "print(0xABCD | 0)\n",
        "43981"
    },
}
