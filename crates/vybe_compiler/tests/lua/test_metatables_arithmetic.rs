//! Arithmetic metamethods: __add, __sub, __mul, __div, __mod, __pow, __unm, __idiv (Lua 5.x §2.4)

lua_print! {
    add_metamethod_op => {
        "local mt={__add=function(a,b) return setmetatable({v=a.v+b.v},mt) end}\nmt.__index=mt\nlocal W=function(n) return setmetatable({v=n},mt) end\nprint((W(3)+W(4)).v)\n",
        "7"
    },
    sub_metamethod_op => {
        "local mt={__sub=function(a,b) return {v=a.v-b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(10)-W(3)).v)\n",
        "7"
    },
    mul_metamethod_op => {
        "local mt={__mul=function(a,b) return {v=a.v*b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(6)*W(7)).v)\n",
        "42"
    },
    div_metamethod_op => {
        "local mt={__div=function(a,b) return {v=a.v/b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(10)/W(4)).v)\n",
        "2.5"
    },
    mod_metamethod_op => {
        "local mt={__mod=function(a,b) return {v=a.v%b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(10)%W(3)).v)\n",
        "1"
    },
    pow_metamethod_op => {
        "local mt={__pow=function(a,b) return {v=a.v^b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(2)^W(8)).v)\n",
        "256.0"
    },
    unm_metamethod_op => {
        "local mt={__unm=function(a) return {v=-a.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((-W(42)).v)\n",
        "-42"
    },
    idiv_metamethod_op => {
        "local mt={__idiv=function(a,b) return {v=a.v//b.v} end}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(7)//W(2)).v)\n",
        "3"
    },
    add_scalar_op => {
        "local mt={__add=function(a,b)\n  local va = type(a)==\"table\" and a.v or a\n  local vb = type(b)==\"table\" and b.v or b\n  return {v=va+vb}\nend}\nlocal W=function(n) return setmetatable({v=n}, mt) end\nprint((W(5)+10).v)\n",
        "15"
    },
    chained_add_op => {
        "local mt={}\nmt.__index=mt\nmt.__add=function(a,b) return setmetatable({v=a.v+b.v},mt) end\nlocal W=function(n) return setmetatable({v=n},mt) end\nprint(((W(1)+W(2))+W(3)).v)\n",
        "6"
    },
}
