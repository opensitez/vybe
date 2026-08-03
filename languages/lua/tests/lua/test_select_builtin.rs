//! `select` built-in — counting varargs and selecting by index (Lua 5.x §6.1)

lua_print! {
    select_count_empty => {
        "local function f(...) return select('#', ...) end\nprint(f())\n",
        "0"
    },
    select_count_three => {
        "local function f(...) return select('#', ...) end\nprint(f(10, 20, 30))\n",
        "3"
    },
    select_count_with_nils => {
        "local function f(...) return select('#', ...) end\nprint(f(1, nil, 3))\n",
        "3"
    },
    select_idx_one_returns_rest => {
        "print(select(1, 'a', 'b', 'c'))\n",
        "a\tb\tc"
    },
    select_idx_two_skips_first => {
        "print(select(2, 'a', 'b', 'c'))\n",
        "b\tc"
    },
    select_idx_three_returns_last => {
        "print(select(3, 'a', 'b', 'c'))\n",
        "c"
    },
    select_negative_index_last => {
        "print(select(-1, 'a', 'b', 'c'))\n",
        "c"
    },
    select_negative_index_second_to_last => {
        "print(select(-2, 'a', 'b', 'c'))\n",
        "b\tc"
    },
    select_used_in_summation => {
        "local function sum(...)\n  local s = 0\n  for i = 1, select('#', ...) do s = s + select(i, ...) end\n  return s\nend\nprint(sum(1, 2, 3, 4))\n",
        "10"
    },
    select_count_nils_only => {
        "local function f(...) return select('#', ...) end\nprint(f(nil, nil))\n",
        "2"
    },
    select_past_bounds_returns_nothing => {
        "local a, b = select(5, 1, 2)\nprint(tostring(a), tostring(b))\n",
        "nil\tnil"
    },
    select_out_of_bounds_negative_raises_error => {
        "local ok = pcall(select, -100, 1, 2)\nprint(ok)\n",
        "false"
    },
    select_zero_index_raises_error => {
        "local ok = pcall(select, 0, 1, 2)\nprint(ok)\n",
        "false"
    },
    select_non_numeric_index_coerced => {
        "print(select('2', 'a', 'b', 'c'))\n",
        "b\tc"
    },
    select_non_coercible_index_raises_error => {
        "local ok = pcall(select, {}, 1, 2)\nprint(ok)\n",
        "false"
    },
    select_float_index_truncated_to_integer => {
        "print(select(2.7, 'a', 'b', 'c'))\n",
        "b\tc"
    },
    select_string_float_index_coerced => {
        "print(select('2.5', 'a', 'b', 'c'))\n",
        "b\tc"
    },
    select_with_no_extra_arguments => {
        "local a, b = select(1)\nprint(tostring(a) .. \",\" .. tostring(b))\n",
        "nil,nil"
    },
    select_hash_with_no_extra_arguments => {
        "print(select('#'))\n",
        "0"
    },
    select_count_argument_is_case_sensitive => {
        "local ok = pcall(select, 'HASH')\nprint(ok)\n",
        "false"
    },
    select_with_negative_one_on_empty_raises_error => {
        "local ok = pcall(select, -1)\nprint(ok)\n",
        "false"
    },
    select_returning_nil_in_middle_of_results => {
        "local a, b, c = select(2, 10, nil, 30)\nprint(tostring(a) .. \",\" .. tostring(b) .. \",\" .. tostring(c))\n",
        "nil,30,nil"
    },
    select_count_with_many_arguments => {
        "print(select('#', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10))\n",
        "10"
    } }
