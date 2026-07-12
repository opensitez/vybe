//! Extended list methods: sort keys, bisect-style insert, comprehensions, index/count, slice assign.

crate::runtime_case!(
    list_sort_key_abs,
    "a = [-3, 1, -2]\na.sort(key=abs)\nprint(a)\n",
    "[1, -2, -3]"
);
crate::runtime_case!(
    list_sort_key_len_strings,
    "a = ['ccc', 'a', 'bb']\na.sort(key=len)\nprint(a)\n",
    "['a', 'bb', 'ccc']"
);
crate::runtime_case!(
    list_sort_reverse_key,
    "a = [3, 1, 2]\na.sort(key=lambda x: -x)\nprint(a)\n",
    "[3, 2, 1]"
);
crate::runtime_case!(
    list_sort_stable_equal_keys,
    "a = [(1, 'b'), (1, 'a'), (2, 'c')]\na.sort(key=lambda t: t[0])\nprint([x[1] for x in a])\n",
    "['b', 'a', 'c']"
);
crate::runtime_case!(
    list_sort_inplace_returns_none,
    "a = [2, 1]\nr = a.sort()\nprint(r, a)\n",
    "None [1, 2]"
);
crate::runtime_case!(
    list_reverse_then_sort,
    "a = [3, 1, 2]\na.reverse()\na.sort()\nprint(a)\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    list_sort_strings_casefold,
    "a = ['B', 'a', 'C']\na.sort(key=str.lower)\nprint(a)\n",
    "['a', 'B', 'C']"
);
crate::runtime_case!(
    list_bisect_insert_sorted_pos,
    "a = [1, 3, 5]\nx = 4\ni = 0\nwhile i < len(a) and a[i] < x:\n    i += 1\na.insert(i, x)\nprint(a)\n",
    "[1, 3, 4, 5]"
);
crate::runtime_case!(
    list_bisect_insert_at_start,
    "a = [2, 4, 6]\nx = 0\ni = 0\nwhile i < len(a) and a[i] < x:\n    i += 1\na.insert(i, x)\nprint(a)\n",
    "[0, 2, 4, 6]"
);
crate::runtime_case!(
    list_bisect_insert_at_end,
    "a = [1, 2, 3]\nx = 9\ni = 0\nwhile i < len(a) and a[i] < x:\n    i += 1\na.insert(i, x)\nprint(a)\n",
    "[1, 2, 3, 9]"
);
crate::runtime_case!(
    list_bisect_insert_duplicate,
    "a = [1, 2, 2, 3]\nx = 2\ni = 0\nwhile i < len(a) and a[i] <= x:\n    i += 1\na.insert(i, x)\nprint(a)\n",
    "[1, 2, 2, 2, 3]"
);
crate::runtime_case!(
    list_comp_nested_squares,
    "print([[i * j for j in range(3)] for i in range(2)])\n",
    "[[0, 0, 0], [0, 1, 2]]"
);
crate::runtime_case!(
    list_comp_if_filter_even,
    "print([x for x in range(6) if x % 2 == 0])\n",
    "[0, 2, 4]"
);
crate::runtime_case!(
    list_comp_if_else_sign,
    "print(['pos' if x > 0 else 'neg' for x in [-1, 0, 2]])\n",
    "['neg', 'neg', 'pos']"
);
crate::runtime_case!(
    list_comp_nested_flatten_manual,
    "print([y for row in [[1, 2], [3]] for y in row])\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    list_comp_zip_pairs,
    "print([a + b for a, b in zip([1, 2], [10, 20])])\n",
    "[11, 22]"
);
crate::runtime_case!(
    list_comp_enumerate_indexed,
    "print([f'{i}:{v}' for i, v in enumerate(['a', 'b'])])\n",
    "['0:a', '1:b']"
);
crate::runtime_case!(
    list_comp_dict_keys_sorted,
    "print([k for k in sorted({'b': 2, 'a': 1})])\n",
    "['a', 'b']"
);
crate::runtime_case!(
    list_comp_set_unique_chars,
    "print(sorted({c for c in 'banana'}))\n",
    "['a', 'b', 'n']"
);
crate::runtime_case!(
    list_comp_range_step,
    "print([x for x in range(0, 10, 3)])\n",
    "[0, 3, 6, 9]"
);
crate::runtime_case!(
    list_comp_double_condition,
    "print([x for x in range(10) if x % 2 == 0 if x > 3])\n",
    "[4, 6, 8]"
);
crate::runtime_case!(
    list_index_first_occurrence,
    "print([1, 2, 1, 3].index(1))\n",
    "0"
);
crate::runtime_case!(
    list_index_middle,
    "print(['a', 'b', 'c', 'd'].index('c'))\n",
    "2"
);
crate::runtime_case!(
    list_index_negative_search,
    "print([10, 20, 30, 40].index(30))\n",
    "2"
);
crate::runtime_case!(list_count_zero, "print([1, 2, 3].count(9))\n", "0");
crate::runtime_case!(list_count_all_same, "print([7, 7, 7].count(7))\n", "3");
crate::runtime_case!(
    list_count_bool_coercion,
    "print([True, 1, False, 0].count(True))\n",
    "1"
);
crate::runtime_case!(
    list_count_sublist_not_found,
    "print([[1], [2]].count([1]))\n",
    "0"
);
crate::runtime_case!(
    list_slice_assign_replace_tail,
    "a = [1, 2, 3, 4]\na[2:] = [30, 40]\nprint(a)\n",
    "[1, 2, 30, 40]"
);
crate::runtime_case!(
    list_slice_assign_replace_head,
    "a = [1, 2, 3, 4]\na[:2] = [10, 20]\nprint(a)\n",
    "[10, 20, 3, 4]"
);
crate::runtime_case!(
    list_slice_assign_step_one_equiv,
    "a = [0, 1, 2, 3]\na[1:3] = [9, 8]\nprint(a)\n",
    "[0, 9, 8, 3]"
);
crate::runtime_case!(
    list_slice_assign_single_to_range,
    "a = [1, 2, 3, 4, 5]\na[1:4] = [99]\nprint(a)\n",
    "[1, 99, 5]"
);
crate::runtime_case!(
    list_slice_assign_empty_deletes,
    "a = [1, 2, 3]\na[1:2] = []\nprint(a)\n",
    "[1, 3]"
);
crate::runtime_case!(
    list_slice_assign_extend_middle,
    "a = [1, 4]\na[1:1] = [2, 3]\nprint(a)\n",
    "[1, 2, 3, 4]"
);
crate::runtime_case!(
    list_slice_assign_full_copy_pattern,
    "a = [1, 2, 3]\nb = a[:]\nb[0] = 9\nprint(a, b)\n",
    "[1, 2, 3] [9, 2, 3]"
);
crate::runtime_case!(
    list_slice_assign_negative_indices,
    "a = [1, 2, 3, 4, 5]\na[-3:-1] = [20, 30]\nprint(a)\n",
    "[1, 2, 20, 30, 5]"
);
crate::runtime_case!(
    list_sort_then_index,
    "a = [30, 10, 20]\na.sort()\nprint(a.index(20))\n",
    "1"
);
crate::runtime_case!(
    list_reverse_copy_independent,
    "a = [1, 2, 3]\nb = a\na.reverse()\nprint(b, a)\n",
    "[3, 2, 1] [3, 2, 1]"
);
crate::runtime_case!(
    list_comp_from_split,
    "print([p for p in 'a,b,c'.split(',')])\n",
    "['a', 'b', 'c']"
);
crate::runtime_case!(
    list_insert_beyond_end_clamps,
    "a = [1, 2]\na.insert(99, 3)\nprint(a)\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    list_sort_mixed_int_compare,
    "a = [3, 1, 2]\na.sort()\nprint(a[0], a[-1])\n",
    "1 3"
);
crate::runtime_case!(
    list_count_after_sort,
    "a = [2, 1, 2, 3]\na.sort()\nprint(a.count(2))\n",
    "2"
);

crate::compile_case!(
    list_sort_reverse_key_compile,
    "a = [3, 1, 2]\na.sort(reverse=True, key=abs)\n"
);
crate::compile_case!(
    list_index_with_start_stop,
    "a = [1, 2, 3, 2, 1]\ni = a.index(2, 2)\n"
);
crate::compile_case!(
    list_remove_all_duplicates_loop,
    "a = [1, 2, 2, 3]\nwhile 2 in a:\n    a.remove(2)\n"
);
crate::compile_case!(list_del_item_by_index, "a = [1, 2, 3]\ndel a[1]\n");
crate::compile_case!(list_del_slice_step, "a = [0, 1, 2, 3, 4, 5]\ndel a[::2]\n");
