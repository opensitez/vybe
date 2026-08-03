//! Advanced variadic function argument forwarding and selecting (Lua 5.x §3.4.11)

lua_print! {
    vararg_forward_print => {
        "local function f(...) print(...) end\nf(1, 2)\n",
        "1	2"
    },
    vararg_select_rest => {
        "local function f(...) return select(2, ...) end\nprint(f(\"a\", \"b\", \"c\"))\n",
        "b\tc"
    },
    vararg_select_count => {
        "local function f(...) return select(\"#\", ...) end\nprint(f(nil, nil, 3))\n",
        "3"
    },
    vararg_pack_n => {
        "local t = table.pack(\"a\", nil, \"c\")\nprint(t.n .. \",\" .. tostring(t[2]))\n",
        "3,nil"
    },
    vararg_sum_recursive => {
        "local function sum(head, ...)\n  if not head then return 0 end\n  return head + sum(...)\nend\nprint(sum(1, 2, 3, 4, 5))\n",
        "15"
    } }
