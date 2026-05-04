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

compile_case!(list_comp_two_fors_compile, "pairs = [(i, j) for i in range(2) for j in range(2)]\n");
compile_case!(list_comp_if_chain_compile, "xs = [x for x in range(10) if x % 2 == 0 if x > 4]\n");
compile_case!(list_comp_nested_ternary_compile, "xs = ['big' if x > 5 else 'small' for x in range(10)]\n");
compile_case!(dict_comp_two_fors_compile, "d = {(i, j): i + j for i in range(2) for j in range(2)}\n");
compile_case!(set_comp_two_fors_compile, "s = {(i, j) for i in range(2) for j in range(2)}\n");
compile_case!(gen_comp_two_fors_compile, "g = ((i, j) for i in range(2) for j in range(2))\n");
compile_case!(comp_with_unpack_compile, "xs = [a + b for a, b in [(1, 2), (3, 4)]]\n");
compile_case!(comp_with_method_call_compile, "xs = [s.strip() for s in [' a ', ' b ']]\n");
compile_case!(comp_with_attr_compile, "xs = [obj.value for obj in items]\n");
compile_case!(comp_with_subscript_compile, "xs = [row[0] for row in [[1, 2], [3, 4]]]\n");
compile_case!(comp_with_walrus_list_compile, "xs = [(y := x * 2) for x in range(3)]\n");
compile_case!(comp_with_walrus_filter_compile, "xs = [y for x in range(5) if (y := x * 2) > 2]\n");
compile_case!(comp_with_walrus_dict_compile, "d = {x: (y := x * 2) for x in range(3)}\n");
compile_case!(comp_with_walrus_set_compile, "s = {(y := x * 2) for x in range(3)}\n");
compile_case!(walrus_in_ifexpr_compile, "x = 1\ny = (z := x + 1) if x else 0\n");
compile_case!(walrus_in_while_compile, "while (line := reader()):\n    print(line)\n");
compile_case!(walrus_in_lambda_compile, "func = lambda x: (y := x + 1)\n");
compile_case!(walrus_nested_compile, "result = [z for x in range(3) if (y := x + 1) and (z := y + 1)]\n");
compile_case!(comp_in_call_compile, "print(sum(x * x for x in range(5)))\n");
compile_case!(comp_in_tuple_compile, "t = tuple(x for x in range(3))\n");
runtime_case!(list_comp_runtime_sum, "print(sum([x * 2 for x in range(4)]))\n", "12");
runtime_case!(dict_comp_runtime_len, "d = {x: x * x for x in range(4)}\nprint(len(d))\n", "4");
runtime_case!(set_comp_runtime_len, "s = {x % 2 for x in range(6)}\nprint(len(s))\n", "2");
runtime_case!(gen_comp_runtime_sum, "print(sum(x for x in range(4)))\n", "6");
compile_case!(nested_comp_matrix_compile, "matrix = [[i * j for j in range(3)] for i in range(3)]\n");
compile_case!(comp_with_zip_compile, "pairs = [(a, b) for a, b in zip([1, 2], [3, 4])]\n");
compile_case!(comp_with_enumerate_compile, "pairs = [(i, x) for i, x in enumerate(['a', 'b'])]\n");
compile_case!(comp_with_sorted_compile, "xs = [x for x in sorted([3, 1, 2])]\n");
compile_case!(comp_with_reversed_compile, "xs = [x for x in reversed([1, 2, 3])]\n");
compile_case!(comp_with_condition_expr_compile, "xs = [x if x > 1 else 0 for x in range(4)]\n");