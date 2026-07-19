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
    set_add_runtime,
    "s = {1, 2}\ns.add(3)\nprint(3 in s)\n",
    "True"
);
runtime_case!(
    set_remove_runtime,
    "s = {1, 2, 3}\ns.remove(2)\nprint(2 in s)\n",
    "False"
);
runtime_case!(
    set_discard_missing_runtime,
    "s = {1, 2}\ns.discard(9)\nprint(len(s))\n",
    "2"
);
runtime_case!(
    set_clear_runtime,
    "s = {1, 2, 3}\ns.clear()\nprint(len(s))\n",
    "0"
);
runtime_case!(
    set_copy_independent_runtime,
    "a = {1, 2}\nb = a.copy()\nb.add(3)\nprint(3 in a)\n",
    "False"
);
runtime_case!(
    set_update_runtime,
    "s = {1}\ns.update({2, 3})\nprint(3 in s)\n",
    "True"
);
runtime_case!(
    set_intersection_update_runtime,
    "s = {1, 2, 3}\ns.intersection_update({2, 3, 4})\nprint(1 in s)\n",
    "False"
);
runtime_case!(
    set_difference_update_runtime,
    "s = {1, 2, 3}\ns.difference_update({2})\nprint(2 in s)\n",
    "False"
);
runtime_case!(
    set_membership_runtime,
    "s = {'a', 'b'}\nprint('b' in s)\n",
    "True"
);
runtime_case!(
    set_union_operator_runtime,
    "a = {1, 2}\nb = {2, 3}\nc = a | b\nprint(3 in c)\n",
    "True"
);
runtime_case!(
    set_intersection_operator_runtime,
    "a = {1, 2}\nb = {2, 3}\nc = a & b\nprint(2 in c)\n",
    "True"
);
runtime_case!(
    set_difference_operator_runtime,
    "a = {1, 2}\nb = {2, 3}\nc = a - b\nprint(1 in c)\n",
    "True"
);
runtime_case!(
    set_symdiff_operator_runtime,
    "a = {1, 2}\nb = {2, 3}\nc = a ^ b\nprint(2 in c)\n",
    "False"
);
runtime_case!(
    set_len_after_duplicate_add_runtime,
    "s = {1}\ns.add(1)\nprint(len(s))\n",
    "1"
);
runtime_case!(
    set_nested_membership_runtime,
    "s = {frozenset({1, 2})}\nprint(frozenset({1, 2}) in s)\n",
    "True"
);

compile_case!(frozenset_constructor_compile, "s = frozenset([1, 2, 3])\n");
compile_case!(
    set_union_method_compile,
    "a = {1, 2}\nb = {2, 3}\nc = a.union(b)\n"
);
compile_case!(
    set_intersection_method_compile,
    "a = {1, 2, 3}\nb = {2, 3, 4}\nc = a.intersection(b)\n"
);
compile_case!(
    set_difference_method_compile,
    "a = {1, 2, 3}\nb = {2}\nc = a.difference(b)\n"
);
compile_case!(
    set_symmetric_difference_method_compile,
    "a = {1, 2}\nb = {2, 3}\nc = a.symmetric_difference(b)\n"
);
compile_case!(
    set_symmetric_difference_update_compile,
    "a = {1, 2}\na.symmetric_difference_update({2, 3})\n"
);
compile_case!(
    set_isdisjoint_compile,
    "a = {1, 2}\nb = {3, 4}\nok = a.isdisjoint(b)\n"
);
compile_case!(
    set_issubset_compile,
    "a = {1, 2}\nb = {1, 2, 3}\nok = a.issubset(b)\n"
);
compile_case!(
    set_issuperset_compile,
    "a = {1, 2, 3}\nb = {1, 2}\nok = a.issuperset(b)\n"
);
compile_case!(set_literal_unpack_compile, "a = {2, 3}\nb = {1, *a, 4}\n");
compile_case!(set_aug_or_compile, "a = {1}\na |= {2, 3}\n");
compile_case!(set_aug_and_compile, "a = {1, 2, 3}\na &= {2, 3, 4}\n");
compile_case!(set_aug_sub_compile, "a = {1, 2, 3}\na -= {2}\n");
compile_case!(set_aug_xor_compile, "a = {1, 2}\na ^= {2, 3}\n");
compile_case!(
    set_comprehension_guard_compile,
    "s = {x * 2 for x in range(10) if x % 2 == 0}\n"
);
