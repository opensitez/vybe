//! Varargs and `select` — Lua 5.x manual §3.4.11.

lua_print! {
    vararg_forwarded_to_print => {
        "function show(...) print(...) end\nshow(1, 2, 3)\n",
        "1	2	3"
    },
    select_picks_value_from_varargs => {
        "print(select(2, \"a\", \"b\", \"c\"))\n",
        "b	c"
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
        "30	40"
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
        "print(select(\"#\", 10, 20, 30))\n",
        "3"
    },
    vararg_in_middle_of_expression_list_adjusted_to_one => {
        "local function f(...) return ..., 99 end\nlocal a, b, c = f(1, 2, 3)\nprint(a, b, c)\n",
        "1\t99\tnil"
    },
    table_pack_preserves_nil_count => {
        "local t = table.pack(10, nil, 30)\nprint(t.n .. ',' .. tostring(t[2]))\n",
        "3,nil"
    },
    vararg_passed_through_two_function_calls => {
        "local function inner(...) return ... end\nlocal function outer(...) return inner(...) end\nprint(outer(5, 6, 7))\n",
        "5\t6\t7"
    },
    vararg_with_select_hash_counts_nils => {
        "local function count(...) return select('#', ...) end\nprint(count(1, nil, 3))\n",
        "3"
    },
    table_unpack_of_packed_varargs_round_trips => {
        "local function pack_and_unpack(...)\n  local t = table.pack(...)\n  return table.unpack(t, 1, t.n)\nend\nlocal a, b, c = pack_and_unpack(7, 8, 9)\nprint(a .. ',' .. b .. ',' .. c)\n",
        "7,8,9"
    },
    vararg_not_visible_in_nested_regular_function => {
        "local function outer(...)\n  local function inner() return select('#', ...) end\n  return inner()\nend\nprint(outer(1, 2, 3))\n",
        "0"
    },
    vararg_used_in_string_format => {
        "local function fmt(pattern, ...)\n  return string.format(pattern, ...)\nend\nprint(fmt('%d + %d = %d', 1, 2, 3))\n",
        "1 + 2 = 3"
    },
    vararg_with_extra_args_past_select_index => {
        "local function from_second(...)\n  return select(2, ...)\nend\nlocal a, b = from_second(10, 20, 30)\nprint(a .. ',' .. b)\n",
        "20,30"
    },
}
