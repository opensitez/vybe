use crate::helpers::{run_print, run_python_one};

#[test]
fn frozenset_literal_from_list() {
    assert_eq!(run_print("frozenset([1, 2, 2])"), "frozenset({1, 2})");
}

#[test]
fn frozenset_empty() {
    assert_eq!(run_print("frozenset()"), "frozenset()");
}

#[test]
fn frozenset_from_string_unique_chars() {
    assert_eq!(run_print("frozenset('aba')"), "frozenset({'a', 'b'})");
}

#[test]
fn frozenset_union_with_set() {
    assert_eq!(
        run_print("frozenset([1, 2]) | {2, 3}"),
        "frozenset({1, 2, 3})"
    );
}

#[test]
fn frozenset_intersection_with_set() {
    assert_eq!(run_print("frozenset([1, 2, 3]) & {2, 4}"), "frozenset({2})");
}

#[test]
fn frozenset_difference() {
    assert_eq!(run_print("frozenset([1, 2, 3]) - {2}"), "frozenset({1, 3})");
}

#[test]
fn frozenset_symmetric_difference() {
    assert_eq!(run_print("frozenset([1, 2]) ^ {2, 3}"), "frozenset({1, 3})");
}

#[test]
fn frozenset_issubset() {
    assert_eq!(run_print("frozenset([1, 2]).issubset({1, 2, 3})"), "True");
}

#[test]
fn frozenset_issuperset() {
    assert_eq!(run_print("frozenset([1, 2, 3]).issuperset({1, 2})"), "True");
}

#[test]
fn frozenset_isdisjoint_true() {
    assert_eq!(run_print("frozenset([1]).isdisjoint({2})"), "True");
}

#[test]
fn frozenset_isdisjoint_false() {
    assert_eq!(run_print("frozenset([1, 2]).isdisjoint({2, 3})"), "False");
}

#[test]
fn frozenset_contains() {
    assert_eq!(run_print("2 in frozenset([1, 2, 3])"), "True");
}

#[test]
fn frozenset_len() {
    assert_eq!(run_print("len(frozenset([1, 1, 2]))"), "2");
}

#[test]
fn frozenset_as_dict_key() {
    assert_eq!(
        run_python_one("fs = frozenset([1])\nd = {fs: 'ok'}\nprint(d[fs])\n"),
        "ok"
    );
}

#[test]
fn frozenset_in_set_of_frozensets() {
    assert_eq!(
        run_python_one("a = frozenset([1])\nb = frozenset([2])\ns = {a, b}\nprint(len(s))\n"),
        "2"
    );
}

#[test]
fn frozenset_copy_returns_self() {
    assert_eq!(
        run_python_one("fs = frozenset([1, 2])\nprint(fs.copy() == fs)\n"),
        "True"
    );
}

#[test]
fn frozenset_equality_with_set_same_elements() {
    assert_eq!(run_print("frozenset([1, 2]) == {1, 2}"), "True");
}

#[test]
fn frozenset_inequality_different_size() {
    assert_eq!(run_print("frozenset([1]) != frozenset([1, 2])"), "True");
}

#[test]
fn frozenset_iter_sorted_list() {
    assert_eq!(run_print("sorted(frozenset([3, 1, 2]))"), "[1, 2, 3]");
}

#[test]
fn frozenset_no_add_method() {
    assert_eq!(run_print("hasattr(frozenset(), 'add')"), "False");
}

#[test]
fn frozenset_no_discard_method() {
    assert_eq!(run_print("hasattr(frozenset(), 'discard')"), "False");
}

#[test]
fn frozenset_has_union_method() {
    assert_eq!(run_print("hasattr(frozenset(), 'union')"), "True");
}

#[test]
fn frozenset_union_method() {
    assert_eq!(
        run_print("frozenset([1]).union([2, 3])"),
        "frozenset({1, 2, 3})"
    );
}

#[test]
fn frozenset_intersection_method() {
    assert_eq!(
        run_print("frozenset([1, 2, 3]).intersection([2, 4])"),
        "frozenset({2})"
    );
}

#[test]
fn frozenset_bool_nonempty_true() {
    assert_eq!(run_print("bool(frozenset([0]))"), "True");
}

#[test]
fn frozenset_bool_empty_false() {
    assert_eq!(run_print("bool(frozenset())"), "False");
}

#[test]
fn frozenset_from_tuple() {
    assert_eq!(run_print("frozenset((1, 2, 1))"), "frozenset({1, 2})");
}

#[test]
fn frozenset_repr_stable() {
    assert_eq!(
        run_python_one("fs = frozenset([1])\nprint(str(fs) == str(fs))\n"),
        "True"
    );
}

#[test]
fn set_of_frozenset_and_set_rejected_if_mixed_unhashable() {
    assert_eq!(
        run_python_one(
            "try:\n {frozenset([1]), [2]}\n print('ok')\nexcept TypeError:\n print('TypeError')\n"
        ),
        "TypeError"
    );
}

#[test]
fn frozenset_nested_in_tuple() {
    assert_eq!(run_print("(frozenset([1]), 2)"), "(frozenset({1}), 2)");
}
