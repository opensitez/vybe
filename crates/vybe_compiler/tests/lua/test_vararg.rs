//! Varargs and `select` — Lua 5.x manual §3.4.11.

lua_print! {
    vararg_forwarded_to_print => {
        "function show(...) print(...) end\nshow(1, 2, 3)\n",
        "1 2 3"
    },
    select_picks_value_from_varargs => {
        "print(select(2, \"a\", \"b\", \"c\"))\n",
        "b"
    },
    varargs_packed_into_table_spread => {
        "function all(...) return {...} end\nlocal t = all(4, 5)\nprint(t[1] + t[2])\n",
        "9"
    },
    select_hash_counts_varargs_after_other_args => {
        "print(select(\"#\", 1, 2, 3, 4))\n",
        "4"
    },
    vararg_count_via_select_hash => {
        "function n(...) return select(\"#\", ...) end\nprint(n(\"a\", \"b\", \"c\"))\n",
        "3"
    },
    select_returns_nth_vararg => {
        "function third(...) return select(3, ...) end\nprint(third(10, 20, 30, 40))\n",
        "30"
    },
    select_negative_index_counts_from_end => {
        "function last(...) return select(-1, ...) end\nprint(last(1, 2, 9))\n",
        "9"
    },
    empty_vararg_list_has_zero_count => {
        "function n(...) return select(\"#\", ...) end\nprint(n())\n",
        "0"
    },
    vararg_captured_in_table_for_closure => {
        "function make(...)\n  local args = {...}\n  return function() return #args end\nend\nprint(make(1, 2, 3)())\n",
        "3"
    },
    ellipsis_in_table_constructor_captures_varargs => {
        "function pack(...)\n  return {...}\nend\nprint(pack(1, 2, 3)[2])\n",
        "2"
    },
    select_zero_returns_total_count => {
        "print(select(0, 10, 20, 30))\n",
        "3"
    },
}
