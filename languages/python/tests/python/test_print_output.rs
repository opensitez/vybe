//! print() sep/end/multiple-args and related stdout semantics.

crate::runtime_case!(print_single_int, "print(42)\n", "42");
crate::runtime_case!(print_single_string, "print('hi')\n", "hi");
crate::runtime_case!(print_two_args_space_sep, "print(1, 2)\n", "1 2");
crate::runtime_case!(print_three_args, "print('a', 'b', 'c')\n", "a b c");
crate::runtime_case!(print_mixed_types, "print(1, 'x', True)\n", "1 x True");
crate::runtime_case!(print_empty_call, "print()\n", "");
crate::runtime_case!(print_none_value, "print(None)\n", "None");
crate::runtime_case!(print_bool_false, "print(False)\n", "False");
crate::runtime_case!(print_bool_true, "print(True)\n", "True");
crate::runtime_case!(print_list_repr, "print([1, 2])\n", "[1, 2]");
crate::runtime_case!(print_dict_repr, "print({'a': 1})\n", "{'a': 1}");
crate::runtime_case!(print_tuple_repr, "print((1, 2))\n", "(1, 2)");
crate::runtime_case!(print_set_repr, "print({1, 2})\n", "{1, 2}");
crate::runtime_case!(print_sep_comma, "print(1, 2, 3, sep=',')\n", "1,2,3");
crate::runtime_case!(print_sep_empty, "print('a', 'b', sep='')\n", "ab");
crate::runtime_case!(print_sep_newline, "print(1, 2, sep='\\n')\n", "1\n2");
crate::runtime_case!(print_sep_colon, "print('x', 'y', sep=':')\n", "x:y");
crate::runtime_case!(
    print_end_no_newline,
    "print('a', end='')\nprint('b')\n",
    "ab"
);
crate::runtime_case!(print_end_dash, "print(1, end='-')\nprint(2)\n", "1-2");
crate::runtime_case!(
    print_end_then_explicit_newline,
    "print('line', end='')\nprint()\n",
    "line"
);
// `run_python_one` joins ALL output lines with "\n" (see helpers.rs), so two
// `print` calls yield "1\n2" — the previous expectation of just "1" could not
// hold. Every other case in this file is single-line, which is why it was the
// only one affected.
crate::runtime_case!(print_first_of_sequence, "print(1)\nprint(2)\n", "1\n2");
crate::runtime_case!(
    print_chained_end_sep,
    "print(1, 2, sep='|', end='!')\nprint(3)\n",
    "1|2!3"
);
crate::runtime_case!(print_float_repr, "print(3.14)\n", "3.14");
crate::runtime_case!(print_negative_zero, "print(-0.0)\n", "-0.0");
crate::runtime_case!(
    print_large_int,
    "print(10 ** 20)\n",
    "100000000000000000000"
);
crate::runtime_case!(
    print_hex_in_fstring_not_print,
    "print(f'{255:#x}')\n",
    "0xff"
);
crate::runtime_case!(print_repr_escape, "print(repr('\\n'))\n", "'\\n'");
crate::runtime_case!(print_star_unpack_list, "print(*[1, 2, 3])\n", "1 2 3");
crate::runtime_case!(print_star_unpack_tuple, "print(*(10, 20))\n", "10 20");
crate::runtime_case!(print_star_with_sep, "print(*['a', 'b'], sep='-')\n", "a-b");
crate::runtime_case!(print_expression_result, "print(2 + 3 * 4)\n", "14");
crate::runtime_case!(print_comparison_result, "print(3 < 5)\n", "True");
crate::runtime_case!(print_len_result, "print(len('abc'))\n", "3");
crate::runtime_case!(print_type_name, "print(type(1).__name__)\n", "int");
crate::runtime_case!(print_nested_list, "print([[1], [2, 3]])\n", "[[1], [2, 3]]");
crate::runtime_case!(print_bytes_repr, "print(b'hi')\n", "b'hi'");
crate::runtime_case!(
    print_two_prints_concat_end,
    "print('no', end='')\nprint('space')\n",
    "nospace"
);
crate::runtime_case!(
    print_zero_args_with_end,
    "print(end='>')\nprint('<')\n",
    "><"
);
crate::runtime_case!(print_sep_only_one_arg, "print(99, sep='?')\n", "99");
crate::runtime_case!(print_multiple_none, "print(None, None)\n", "None None");
crate::runtime_case!(print_range_object, "print(range(3))\n", "range(0, 3)");
crate::runtime_case!(
    print_enumerate_lazy,
    "print(list(enumerate(['a'])))\n",
    "[(0, 'a')]"
);
crate::runtime_case!(
    print_zip_pairs,
    "print(list(zip([1, 2], ['a', 'b'])))\n",
    "[(1, 'a'), (2, 'b')]"
);
crate::runtime_case!(
    print_sorted_keys,
    "print(sorted({'b': 2, 'a': 1}))\n",
    "['a', 'b']"
);
crate::runtime_case!(print_joined_string, "print('-'.join(['x', 'y']))\n", "x-y");

crate::compile_case!(print_flush_kwarg, "print('x', flush=True)\n");
crate::compile_case!(
    print_file_kwarg,
    "import sys\nprint('x', file=sys.stdout)\n"
);
crate::compile_case!(print_sep_end_together, "print(1, 2, sep=':', end=';')\n");
