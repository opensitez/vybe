//! Vararg `...` expression in constructor and select context (Lua 5.x §3.4.11)

lua_print! {
    vararg_constructor => {
        "local function f(...) return {...} end\nlocal t = f(1, 2, 3)\nprint(#t)\n",
        "3"
    },
    vararg_index => {
        "local function f(...) local t = {...}; return t[2] end\nprint(f(10, 20, 30))\n",
        "20"
    },
    vararg_table_pack_n => {
        "local function f(...)\n  local t = table.pack(...)\n  return t.n\nend\nprint(f(5, nil, 7))\n",
        "3"
    },
    vararg_forward => {
        "local function add(a, b) return a + b end\nlocal function proxy(...) return add(...) end\nprint(proxy(3, 4))\n",
        "7"
    },
    vararg_select_sum => {
        "local function sum(...)\n  local s = 0\n  for i = 1, select('#', ...) do\n    s = s + select(i, ...)\n  end\n  return s\nend\nprint(sum(1, 2, 3, 4, 5))\n",
        "15"
    },
    vararg_non_tail_truncate => {
        "local function f(...) return ... end\nlocal a, b = f(10, 20), 99\nprint(a, b)\n",
        "10\t99"
    },
    vararg_format => {
        "local function fmt(pattern, ...)\n  return string.format(pattern, ...)\nend\nprint(fmt(\"%d+%d=%d\", 2, 3, 5))\n",
        "2+3=5"
    },
    vararg_select_count_nils => {
        "local function f(...) return select('#', ...) end\nprint(f(1, nil, 3))\n",
        "3"
    },
    vararg_select_count_empty => {
        "local function f(...) return select('#', ...) end\nprint(f())\n",
        "0"
    },
    vararg_to_pcall => {
        "local function f(a, b) return a + b end\nlocal ok, v = pcall(f, 7, 8)\nprint(ok, v)\n",
        "true\t15"
    },
    vararg_recursive => {
        "local function concat(...)\n  if select('#', ...) == 0 then return \"\" end\n  return tostring((select(1, ...))) .. concat(select(2, ...))\nend\nprint(concat(1, 2, 3))\n",
        "123"
    },
}
