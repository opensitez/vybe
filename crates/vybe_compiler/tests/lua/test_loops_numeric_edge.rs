//! Numeric `for` loop edge cases — float step, negative step, single iteration (Lua 5.x §3.3.5)

lua_print! {
    numeric_for_float => {
        "local s = 0.0\nfor i = 0.0, 1.0, 0.5 do s = s + i end\nprint(s)\n",
        "1.5"
    },
    numeric_for_neg_step => {
        "local s = 0\nfor i = 5, 1, -1 do s = s + i end\nprint(s)\n",
        "15"
    },
    numeric_for_single_iter => {
        "local n = 0\nfor i = 3, 3 do n = n + 1 end\nprint(n)\n",
        "1"
    },
    numeric_for_zero_iter => {
        "local n = 0\nfor i = 5, 1 do n = n + 1 end\nprint(n)\n",
        "0"
    },
    numeric_for_zero_iter_neg => {
        "local n = 0\nfor i = 1, 5, -1 do n = n + 1 end\nprint(n)\n",
        "0"
    },
    numeric_for_local_scope => {
        "local i = 99\nfor i = 1, 3 do end\nprint(i)\n",
        "99"
    },
    numeric_for_break => {
        "local last = 0\nfor i = 1, 10 do\n  if i == 5 then break end\n  last = i\nend\nprint(last)\n",
        "4"
    },
    numeric_for_large_step => {
        "local s = 0\nfor i = 1, 10, 3 do s = s + i end\nprint(s)\n",
        "22"
    },
    numeric_for_product => {
        "local p = 1\nfor i = 1, 5 do p = p * i end\nprint(p)\n",
        "120"
    },
    numeric_for_limit_once => {
        "local calls = 0\nlocal function limit() calls = calls + 1; return 3 end\nfor i = 1, limit() do end\nprint(calls)\n",
        "1"
    },
    numeric_for_step_once => {
        "local calls = 0\nlocal function step() calls = calls + 1; return 1 end\nfor i = 1, 3, step() do end\nprint(calls)\n",
        "1"
    },
}
