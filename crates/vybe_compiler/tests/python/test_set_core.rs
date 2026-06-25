use crate::helpers::{run_python_one, run_print};

#[test]
fn set_literal_unique() {
    assert_eq!(run_print("sorted({3, 1, 2, 1})"), "[1, 2, 3]");
}

#[test]
fn set_empty() {
    assert_eq!(run_print("set()"), "set()");
}

#[test]
fn set_add_member() {
    assert_eq!(
        run_python_one("s = {1}\ns.add(2)\nprint(sorted(s))\n"),
        "[1, 2]"
    );
}

#[test]
fn set_remove_existing() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns.remove(1)\nprint(sorted(s))\n"),
        "[2]"
    );
}

#[test]
fn set_discard_missing() {
    assert_eq!(
        run_python_one("s = {1}\ns.discard(9)\nprint(sorted(s))\n"),
        "[1]"
    );
}

#[test]
fn set_pop_arbitrary() {
    assert_eq!(
        run_python_one("s = {9}\nprint(s.pop())\n"),
        "9"
    );
}

#[test]
fn set_clear() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns.clear()\nprint(s)\n"),
        "set()"
    );
}

#[test]
fn set_union_operator() {
    assert_eq!(run_print("sorted({1, 2} | {2, 3})"), "[1, 2, 3]");
}

#[test]
fn set_intersection_operator() {
    assert_eq!(run_print("sorted({1, 2, 3} & {2, 3, 4})"), "[2, 3]");
}

#[test]
fn set_difference_operator() {
    assert_eq!(run_print("sorted({1, 2, 3} - {2})"), "[1, 3]");
}

#[test]
fn set_symmetric_difference_operator() {
    assert_eq!(run_print("sorted({1, 2} ^ {2, 3})"), "[1, 3]");
}

#[test]
fn set_subset_operator() {
    assert_eq!(run_print("{1, 2} <= {1, 2, 3}"), "True");
}

#[test]
fn set_proper_subset() {
    assert_eq!(run_print("{1} < {1, 2}"), "True");
}

#[test]
fn set_superset_operator() {
    assert_eq!(run_print("{1, 2, 3} >= {1, 2}"), "True");
}

#[test]
fn set_proper_superset() {
    assert_eq!(run_print("{1, 2, 3} > {1, 2}"), "True");
}

#[test]
fn set_disjoint() {
    assert_eq!(run_print("{1, 2}.isdisjoint({3, 4})"), "True");
}

#[test]
fn set_not_disjoint() {
    assert_eq!(run_print("{1, 2}.isdisjoint({2, 3})"), "False");
}

#[test]
fn set_in_membership() {
    assert_eq!(run_print("2 in {1, 2, 3}"), "True");
}

#[test]
fn set_len() {
    assert_eq!(run_print("len({1, 2, 3})"), "3");
}

#[test]
fn set_bool_empty_false() {
    assert_eq!(run_print("bool(set())"), "False");
}

#[test]
fn set_bool_nonempty_true() {
    assert_eq!(run_print("bool({0})"), "True");
}

#[test]
fn set_from_list() {
    assert_eq!(run_print("sorted(set([1, 1, 2]))"), "[1, 2]");
}

#[test]
fn set_from_string() {
    assert_eq!(run_print("sorted(set('aba'))"), "['a', 'b']");
}

#[test]
fn set_copy() {
    assert_eq!(
        run_python_one("a = {1, 2}\nb = a.copy()\nb.add(3)\nprint(sorted(a), sorted(b))\n"),
        "[1, 2] [1, 2, 3]"
    );
}

#[test]
fn set_update_method() {
    assert_eq!(
        run_python_one("s = {1}\ns.update({2, 3})\nprint(sorted(s))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn set_intersection_update() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns &= {2, 3, 4}\nprint(sorted(s))\n"),
        "[2, 3]"
    );
}

#[test]
fn set_union_update() {
    assert_eq!(
        run_python_one("s = {1}\ns |= {2}\nprint(sorted(s))\n"),
        "[1, 2]"
    );
}

#[test]
fn set_difference_update() {
    assert_eq!(
        run_python_one("s = {1, 2, 3}\ns -= {2}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn set_symmetric_difference_update() {
    assert_eq!(
        run_python_one("s = {1, 2}\ns ^= {2, 3}\nprint(sorted(s))\n"),
        "[1, 3]"
    );
}

#[test]
fn set_equality() {
    assert_eq!(run_print("{1, 2} == {2, 1}"), "True");
}

#[test]
fn set_inequality() {
    assert_eq!(run_print("{1} == {1, 2}"), "False");
}

#[test]
fn set_frozenset_hashable() {
    assert_eq!(run_print("len({frozenset({1}), frozenset({1})})"), "1");
}

#[test]
fn set_mutable_not_hashable() {
    assert_eq!(
        run_python_one("try:\n hash({1})\nexcept TypeError:\n print('no')\n"),
        "no"
    );
}

#[test]
fn set_iter_order_insertion() {
    assert_eq!(
        run_python_one("s = set()\nfor x in [3, 1, 2]:\n s.add(x)\nprint(list(s))\n"),
        "[3, 1, 2]"
    );
}

#[test]
fn set_comprehension_inline() {
    assert_eq!(run_print("sorted({x for x in range(3)})"), "[0, 1, 2]");
}

#[test]
fn set_unpack_in_literal() {
    assert_eq!(run_print("sorted({*{1, 2}, *{2, 3}})"), "[1, 2, 3]");
}

#[test]
fn set_remove_raises_keyerror() {
    assert_eq!(
        run_python_one("try:\n {1}.remove(2)\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn set_pop_empty_raises() {
    assert_eq!(
        run_python_one("try:\n set().pop()\nexcept KeyError:\n print('empty')\n"),
        "empty"
    );
}

#[test]
fn set_issubset_method() {
    assert_eq!(run_print("{1, 2}.issubset({1, 2, 3})"), "True");
}

#[test]
fn set_issuperset_method() {
    assert_eq!(run_print("{1, 2, 3}.issuperset({1})"), "True");
}

#[test]
fn set_intersection_method() {
    assert_eq!(run_print("sorted({1, 2}.intersection({2, 3}))"), "[2]");
}

#[test]
fn set_union_method() {
    assert_eq!(run_print("sorted({1}.union({2}))"), "[1, 2]");
}

#[test]
fn set_difference_method() {
    assert_eq!(run_print("sorted({1, 2}.difference({2}))"), "[1]");
}

#[test]
fn set_symmetric_difference_method() {
    assert_eq!(run_print("sorted({1, 2}.symmetric_difference({2, 3}))"), "[1, 3]");
}

#[test]
fn set_any_all() {
    assert_eq!(run_print("[any({0, 1}), all({1, 2})]"), "[True, True]");
}
