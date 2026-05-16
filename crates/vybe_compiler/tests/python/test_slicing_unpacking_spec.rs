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

runtime_case!(slice_open_left_runtime, "x = [1, 2, 3, 4]\nprint(x[:2][1])\n", "2");
runtime_case!(slice_open_right_runtime, "x = [1, 2, 3, 4]\nprint(x[2:][0])\n", "3");
runtime_case!(slice_out_of_bounds_runtime, "x = [1, 2, 3]\nprint(len(x[:99]))\n", "3");
runtime_case!(slice_negative_step_runtime, "x = [1, 2, 3, 4]\nprint(x[::-1][0])\n", "4");
runtime_case!(string_reverse_runtime, "s = 'stressed'\nprint(s[::-1])\n", "desserts");
runtime_case!(string_stride_runtime, "s = 'abcdef'\nprint(s[::2])\n", "ace");
compile_case!(tuple_slice_compile, "t = (1, 2, 3, 4)\nx = t[1:3]\n");
compile_case!(bytes_slice_compile, "b = b'abcdef'\nx = b[1:4]\n");
runtime_case!(slice_assign_expand_runtime, "x = [1, 4]\nx[1:1] = [2, 3]\nprint(x[2])\n", "3");
runtime_case!(slice_assign_shrink_runtime, "x = [1, 2, 3, 4]\nx[1:3] = [9]\nprint(x[1])\n", "9");
runtime_case!(slice_assign_replace_all_runtime, "x = [1, 2, 3]\nx[:] = [7, 8]\nprint(len(x))\n", "2");
compile_case!(slice_assign_step_compile, "x = [0, 1, 2, 3, 4, 5]\nx[::2] = [9, 9, 9]\n");
runtime_case!(del_slice_runtime, "x = [1, 2, 3, 4]\ndel x[1:3]\nprint(len(x))\n", "2");
compile_case!(del_step_slice_compile, "x = [1, 2, 3, 4, 5, 6]\ndel x[::2]\n");
compile_case!(starred_head_tail_compile, "first, *middle, last = [1, 2, 3, 4, 5]\n");
compile_case!(starred_ignore_head_compile, "*_, last = [1, 2, 3]\n");
compile_case!(starred_only_middle_compile, "head, *body = [1, 2, 3, 4]\n");
compile_case!(nested_starred_compile, "(a, *b), c = (1, 2, 3), 4\n");
compile_case!(list_literal_star_compile, "a = [2, 3]\nb = [1, *a, 4]\n");
compile_case!(tuple_literal_star_compile, "a = (2, 3)\nb = (1, *a, 4)\n");
compile_case!(set_literal_star_compile, "a = {2, 3}\nb = {1, *a, 4}\n");
compile_case!(dict_literal_star_compile, "a = {'x': 1}\nb = {'y': 2}\nc = {**a, **b}\n");
compile_case!(call_multiple_star_compile, "def f(a, b, c, d):\n    pass\nleft = [1, 2]\nright = [3, 4]\nf(*left, *right)\n");
compile_case!(call_star_and_doublestar_compile, "def f(a, b, c):\n    pass\nargs = [1]\nkw = {'b': 2, 'c': 3}\nf(*args, **kw)\n");
runtime_case!(swap_unpack_runtime, "a, b = 1, 2\na, b = b, a\nprint(a)\n", "2");
runtime_case!(parallel_unpack_runtime, "a, b, c = [1, 2, 3]\nprint(c)\n", "3");
compile_case!(nested_unpack_compile, "[a, (b, c)] = [1, (2, 3)]\n");
compile_case!(slice_object_compile, "sl = slice(1, 5, 2)\n");
compile_case!(subscript_with_slice_object_compile, "x = [1, 2, 3, 4, 5]\nsl = slice(1, 4)\ny = x[sl]\n");
compile_case!(unpack_generator_compile, "def gen():\n    yield 1\n    yield 2\na, b = gen()\n");