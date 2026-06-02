use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

runtime_case!(
    list_extend_runtime,
    "x = [1]\nx.extend([2, 3])\nprint(x[2])\n",
    "3"
);
runtime_case!(
    list_insert_runtime,
    "x = [1, 3]\nx.insert(1, 2)\nprint(x[1])\n",
    "2"
);
runtime_case!(
    list_remove_runtime,
    "x = [1, 2, 3]\nx.remove(2)\nprint(len(x))\n",
    "2"
);
runtime_case!(
    list_pop_last_runtime,
    "x = [1, 2, 3]\nprint(x.pop())\n",
    "3"
);
runtime_case!(
    list_pop_index_runtime,
    "x = [10, 20, 30]\nprint(x.pop(1))\n",
    "20"
);
runtime_case!(
    list_clear_runtime,
    "x = [1, 2, 3]\nx.clear()\nprint(len(x))\n",
    "0"
);
runtime_case!(
    list_copy_independent_runtime,
    "x = [1, 2]\ny = x.copy()\ny.append(3)\nprint(len(x))\n",
    "2"
);
runtime_case!(
    list_reverse_runtime,
    "x = [1, 2, 3]\nx.reverse()\nprint(x[0])\n",
    "3"
);
runtime_case!(
    list_sort_runtime,
    "x = [3, 1, 2]\nx.sort()\nprint(x[0])\n",
    "1"
);
runtime_case!(
    list_count_runtime,
    "x = [1, 2, 2, 3]\nprint(x.count(2))\n",
    "2"
);
runtime_case!(
    list_index_runtime,
    "x = ['a', 'b', 'c']\nprint(x.index('c'))\n",
    "2"
);
runtime_case!(
    list_append_then_len_runtime,
    "x = []\nx.append('a')\nprint(len(x))\n",
    "1"
);
runtime_case!(
    list_extend_then_len_runtime,
    "x = []\nx.extend([1, 2, 3, 4])\nprint(len(x))\n",
    "4"
);
runtime_case!(
    list_slice_assign_expand_runtime,
    "x = [1, 4]\nx[1:1] = [2, 3]\nprint(len(x))\n",
    "4"
);
runtime_case!(
    list_slice_assign_shrink_runtime,
    "x = [1, 2, 3, 4]\nx[1:3] = [9]\nprint(len(x))\n",
    "3"
);
runtime_case!(
    list_slice_replace_all_runtime,
    "x = [1, 2, 3]\nx[:] = [4, 5]\nprint(x[0])\n",
    "4"
);
runtime_case!(
    list_nested_mutation_runtime,
    "x = [[1], [2]]\nx[0].append(9)\nprint(x[0][1])\n",
    "9"
);

compile_case!(
    list_sort_reverse_compile,
    "x = [3, 1, 2]\nx.sort(reverse=True)\n"
);
compile_case!(
    list_sort_key_compile,
    "x = ['bbb', 'a', 'cc']\nx.sort(key=len)\n"
);
compile_case!(
    list_index_start_stop_compile,
    "x = [1, 2, 3, 2]\ni = x.index(2, 2)\n"
);
compile_case!(
    list_insert_negative_compile,
    "x = [1, 2, 3]\nx.insert(-1, 9)\n"
);
compile_case!(
    list_remove_first_duplicate_compile,
    "x = [1, 2, 2, 3]\nx.remove(2)\n"
);
compile_case!(
    list_count_missing_compile,
    "x = [1, 2, 3]\nn = x.count(99)\n"
);
compile_case!(list_extend_tuple_compile, "x = [1]\nx.extend((2, 3))\n");
compile_case!(
    list_slice_assign_step_compile,
    "x = [0, 1, 2, 3, 4, 5]\nx[::2] = [9, 9, 9]\n"
);
compile_case!(
    list_slice_assign_empty_compile,
    "x = [1, 2, 3]\nx[1:2] = []\n"
);
compile_case!(list_del_slice_compile, "x = [1, 2, 3, 4]\ndel x[1:3]\n");
compile_case!(
    list_del_step_slice_compile,
    "x = [1, 2, 3, 4, 5, 6]\ndel x[::2]\n"
);
compile_case!(list_inplace_concat_compile, "x = [1, 2]\nx += [3, 4]\n");
compile_case!(list_inplace_repeat_compile, "x = [1]\nx *= 3\n");
