//! Extended slicing: extended syntax, assignment, deletion, step edge cases.

crate::runtime_case!(slice_basic, "a = [0, 1, 2, 3]\nprint(a[1:3])\n", "[1, 2]");
crate::runtime_case!(
    slice_negative_start,
    "a = [0, 1, 2, 3]\nprint(a[-3:-1])\n",
    "[1, 2]"
);
crate::runtime_case!(
    slice_negative_end,
    "a = [0, 1, 2, 3]\nprint(a[:-1])\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    slice_step_two,
    "a = [0, 1, 2, 3, 4]\nprint(a[::2])\n",
    "[0, 2, 4]"
);
crate::runtime_case!(
    slice_step_negative,
    "a = [0, 1, 2, 3]\nprint(a[::-1])\n",
    "[3, 2, 1, 0]"
);
crate::runtime_case!(slice_empty_range, "a = [1, 2, 3]\nprint(a[2:2])\n", "[]");
crate::runtime_case!(slice_oob_start, "a = [1, 2, 3]\nprint(a[5:])\n", "[]");
crate::runtime_case!(slice_oob_end, "a = [1, 2, 3]\nprint(a[:10])\n", "[1, 2, 3]");
crate::runtime_case!(
    slice_assign_replace,
    "a = [1, 2, 3]\na[1:2] = [9]\nprint(a)\n",
    "[1, 9, 3]"
);
crate::runtime_case!(
    slice_assign_insert,
    "a = [1, 3]\na[1:1] = [2]\nprint(a)\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    slice_assign_delete,
    "a = [1, 2, 3]\na[1:2] = []\nprint(a)\n",
    "[1, 3]"
);
crate::runtime_case!(
    slice_assign_step,
    "a = [0, 1, 2, 3, 4]\na[::2] = [9, 9, 9]\nprint(a)\n",
    "[9, 1, 9, 3, 9]"
);
crate::runtime_case!(
    slice_full_copy,
    "a = [1, 2]\nb = a[:]\nb[0] = 9\nprint(a)\n",
    "[1, 2]"
);
crate::runtime_case!(slice_string, "print('abcdef'[2:5])\n", "cde");
crate::runtime_case!(slice_bytes, "print(b'abcdef'[1:4])\n", "b'bcd'");
crate::runtime_case!(slice_tuple, "print((1, 2, 3, 4)[1:3])\n", "(2, 3)");
crate::runtime_case!(
    slice_range_object,
    "print(list(range(5)[1:4]))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    slice_del_item,
    "a = [1, 2, 3]\ndel a[1]\nprint(a)\n",
    "[1, 3]"
);
crate::runtime_case!(
    slice_del_range,
    "a = [1, 2, 3, 4]\ndel a[1:3]\nprint(a)\n",
    "[1, 4]"
);
crate::runtime_case!(
    slice_del_step,
    "a = [0, 1, 2, 3, 4]\ndel a[::2]\nprint(a)\n",
    "[1, 3]"
);
crate::runtime_case!(
    slice_negative_step,
    "a = [0, 1, 2, 3]\nprint(a[3:0:-1])\n",
    "[3, 2, 1]"
);
crate::runtime_case!(
    slice_assign_negative,
    "a = [1, 2, 3, 4, 5]\na[-3:-1] = [20, 30]\nprint(a)\n",
    "[1, 2, 20, 30, 5]"
);
crate::runtime_case!(
    slice_zero_step_error,
    "a = [1, 2, 3]\ntry:\n a[::0]\n print('ok')\nexcept ValueError:\n print('err')\n",
    "err"
);
crate::runtime_case!(slice_index_single, "a = [10, 20, 30]\nprint(a[-1])\n", "30");
crate::runtime_case!(
    slice_nested_list,
    "a = [[1], [2, 3]]\nprint(a[1][0])\n",
    "2"
);
crate::runtime_case!(
    slice_bytearray,
    "b = bytearray(b'abcd')\nprint(b[1:3])\n",
    "bytearray(b'bc')"
);
crate::runtime_case!(
    slice_assign_bytearray,
    "b = bytearray(b'abcd')\nb[1:3] = b'XY'\nprint(b)\n",
    "bytearray(b'aXYd')"
);
crate::runtime_case!(
    slice_extended_index,
    "a = [0, 1, 2, 3, 4, 5]\nprint(a[1:5:2])\n",
    "[1, 3]"
);
crate::runtime_case!(
    slice_none_bounds,
    "a = [1, 2, 3]\nprint(a[:])\n",
    "[1, 2, 3]"
);
crate::runtime_case!(slice_start_only, "a = [1, 2, 3]\nprint(a[1:])\n", "[2, 3]");
crate::runtime_case!(slice_end_only, "a = [1, 2, 3]\nprint(a[:2])\n", "[1, 2]");
crate::runtime_case!(
    slice_large_step,
    "a = [0, 1, 2, 3]\nprint(a[::10])\n",
    "[0]"
);
crate::runtime_case!(
    slice_reverse_assign,
    "a = [1, 2, 3]\na[::-1] = [4, 5, 6]\nprint(a)\n",
    "[6, 5, 4]"
);
crate::runtime_case!(slice_string_step, "print('abcdef'[::2])\n", "ace");
crate::runtime_case!(slice_tuple_negative, "print((1, 2, 3, 4)[-2:])\n", "(3, 4)");
crate::runtime_case!(
    slice_list_comp,
    "print([x for x in range(5)][1:4])\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    slice_dict_keys_list,
    "print(list({'a': 1, 'b': 2})[0:1])\n",
    "['a']"
);
crate::runtime_case!(slice_empty_list, "print([][:])\n", "[]");
crate::runtime_case!(slice_single_element, "print([99][0:1])\n", "[99]");
crate::runtime_case!(
    slice_assign_shorter,
    "a = [1, 2, 3, 4]\na[1:3] = [9]\nprint(a)\n",
    "[1, 9, 4]"
);
crate::runtime_case!(
    slice_assign_longer,
    "a = [1, 2, 3]\na[1:2] = [20, 30, 40]\nprint(a)\n",
    "[1, 20, 30, 40, 3]"
);
crate::runtime_case!(slice_mod_index, "a = [0, 1, 2]\nprint(a[10:11])\n", "[]");
crate::runtime_case!(slice_on_range, "print(list(range(10)[2:8:3]))\n", "[2, 5]");
crate::runtime_case!(
    slice_memoryview_like,
    "print(bytes([0,1,2,3])[1:3])\n",
    "b'\\x01\\x02'"
);
crate::runtime_case!(
    slice_chained,
    "a = [0, 1, 2, 3, 4, 5]\nprint(a[1:5][1:3])\n",
    "[2, 3]"
);
crate::runtime_case!(slice_bool_context, "print(bool([1][1:1]))\n", "False");
crate::runtime_case!(
    slice_len_after,
    "a = [1, 2, 3, 4]\nprint(len(a[1:3]))\n",
    "2"
);
crate::runtime_case!(slice_equality, "print([1, 2, 3][1:] == [2, 3])\n", "True");

crate::compile_case!(
    slice_assign_negative_step,
    "a = [0,1,2,3,4]\na[4:0:-2] = [9,9]\n"
);
crate::compile_case!(
    slice_object_index,
    "class S:\n def __getitem__(self, i):\n  return i\nS()[1:2]\n"
);
crate::compile_case!(
    slice_extended_tuple_target,
    "a = [1,2,3,4,5,6]\na[1:5:2] = [9,9]\n"
);
crate::compile_case!(slice_del_negative_step, "a = [1,2,3,4]\ndel a[::-2]\n");
crate::compile_case!(
    slice_memoryview,
    "mv = memoryview(b'abcd')\nbytes(mv[1:3])\n"
);
