//! Extended set/frozenset methods: algebra, update variants, discard/remove, comprehension.

crate::runtime_case!(
    set_add_member,
    "s = {1}\ns.add(2)\nprint(sorted(s))\n",
    "[1, 2]"
);
crate::runtime_case!(
    set_remove_existing,
    "s = {1, 2, 3}\ns.remove(2)\nprint(sorted(s))\n",
    "[1, 3]"
);
crate::runtime_case!(
    set_discard_missing,
    "s = {1}\ns.discard(9)\nprint(len(s))\n",
    "1"
);
crate::runtime_case!(
    set_pop_arbitrary,
    "s = {10, 20, 30}\nx = s.pop()\nprint(x in {10, 20, 30})\n",
    "True"
);
crate::runtime_case!(set_clear, "s = {1, 2}\ns.clear()\nprint(len(s))\n", "0");
crate::runtime_case!(
    set_union_pipe,
    "print(sorted({1, 2} | {2, 3}))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    set_intersection_ampersand,
    "print(sorted({1, 2, 3} & {2, 3, 4}))\n",
    "[2, 3]"
);
crate::runtime_case!(
    set_difference_minus,
    "print(sorted({1, 2, 3} - {2}))\n",
    "[1, 3]"
);
crate::runtime_case!(
    set_symmetric_difference_caret,
    "print(sorted({1, 2} ^ {2, 3}))\n",
    "[1, 3]"
);
crate::runtime_case!(set_subset_le, "print({1, 2} <= {1, 2, 3})\n", "True");
crate::runtime_case!(set_proper_subset_lt, "print({1} < {1, 2})\n", "True");
crate::runtime_case!(set_superset_ge, "print({1, 2, 3} >= {1, 2})\n", "True");
crate::runtime_case!(
    set_proper_superset_gt,
    "print({1, 2, 3} > {1, 2})\n",
    "True"
);
crate::runtime_case!(set_disjoint, "print({1, 2}.isdisjoint({3, 4}))\n", "True");
crate::runtime_case!(
    set_not_disjoint,
    "print({1, 2}.isdisjoint({2, 3}))\n",
    "False"
);
crate::runtime_case!(
    set_update_inplace,
    "s = {1}\ns.update({2, 3})\nprint(sorted(s))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    set_intersection_update,
    "s = {1, 2, 3}\ns &= {2, 3, 4}\nprint(sorted(s))\n",
    "[2, 3]"
);
crate::runtime_case!(
    set_difference_update,
    "s = {1, 2, 3}\ns -= {2}\nprint(sorted(s))\n",
    "[1, 3]"
);
crate::runtime_case!(
    set_symmetric_difference_update,
    "s = {1, 2}\ns ^= {2, 3}\nprint(sorted(s))\n",
    "[1, 3]"
);
crate::runtime_case!(
    set_comp_unique_squares,
    "print(sorted({x * x for x in [-2, -1, 0, 1, 2]}))\n",
    "[0, 1, 4]"
);
crate::runtime_case!(
    set_comp_from_string,
    "print(sorted({c.lower() for c in 'Hello'}))\n",
    "['e', 'h', 'l', 'o']"
);
crate::runtime_case!(
    set_comp_filter,
    "print(sorted({x for x in range(6) if x % 2 == 1}))\n",
    "[1, 3, 5]"
);
crate::runtime_case!(set_literal_empty, "print(len(set()))\n", "0");
crate::runtime_case!(
    set_from_list_dedup,
    "print(sorted(set([1, 1, 2, 3, 2])))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(set_membership, "print(2 in {1, 2, 3})\n", "True");
crate::runtime_case!(set_len, "print(len({1, 2, 3}))\n", "3");
crate::runtime_case!(set_bool_nonempty, "print(bool({0}))\n", "True");
crate::runtime_case!(set_bool_empty, "print(bool(set()))\n", "False");
crate::runtime_case!(
    frozenset_hashable_in_set,
    "print(len({frozenset({1, 2}), frozenset({1, 2})}))\n",
    "1"
);
crate::runtime_case!(
    frozenset_union,
    "print(sorted(frozenset({1, 2}) | frozenset({2, 3})))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    set_copy_independent,
    "a = {1, 2}\nb = a.copy()\nb.add(3)\nprint(len(a), len(b))\n",
    "2 3"
);
crate::runtime_case!(set_equality, "print({1, 2} == {2, 1})\n", "True");
crate::runtime_case!(set_inequality, "print({1, 2} == {1, 2, 3})\n", "False");
crate::runtime_case!(
    set_union_method,
    "print(sorted({1}.union({2, 3})))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    set_intersection_method,
    "print(sorted({1, 2, 3}.intersection({2, 3, 4})))\n",
    "[2, 3]"
);
crate::runtime_case!(
    set_difference_method,
    "print(sorted({1, 2, 3}.difference({3})))\n",
    "[1, 2]"
);
crate::runtime_case!(
    set_symmetric_difference_method,
    "print(sorted({1, 2}.symmetric_difference({2, 3})))\n",
    "[1, 3]"
);
crate::runtime_case!(set_issubset_method, "print({1}.issubset({1, 2}))\n", "True");
crate::runtime_case!(
    set_issuperset_method,
    "print({1, 2, 3}.issuperset({1, 2}))\n",
    "True"
);
crate::runtime_case!(
    set_unpack_in_literal,
    "print(sorted({*{1, 2}, 3}))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    set_comp_nested_gen,
    "print(sorted({(x, y) for x in range(2) for y in range(2)}))\n",
    "[(0, 0), (0, 1), (1, 0), (1, 1)]"
);
crate::runtime_case!(set_of_tuples_len, "print(len({(1, 2), (3, 4)}))\n", "2");
crate::runtime_case!(
    set_remove_then_len,
    "s = {1, 2, 3, 4}\ns.remove(4)\nprint(len(s))\n",
    "3"
);
crate::runtime_case!(
    set_pop_empty_raises,
    "s = set()\ntry:\n    s.pop()\n    print('ok')\nexcept KeyError:\n    print('empty')\n",
    "empty"
);

crate::compile_case!(
    set_intersection_update_method,
    "s = {1, 2}\ns.intersection_update({2, 3})\n"
);
crate::compile_case!(
    set_difference_update_method,
    "s = {1, 2}\ns.difference_update({1})\n"
);
crate::compile_case!(
    set_symmetric_difference_update_method,
    "s = {1}\ns.symmetric_difference_update({1, 2})\n"
);
crate::compile_case!(set_update_multiple, "s = {1}\ns.update({2}, {3})\n");
crate::compile_case!(frozenset_copy, "f = frozenset({1, 2})\ng = f.copy()\n");
