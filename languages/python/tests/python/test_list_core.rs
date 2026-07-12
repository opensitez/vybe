use crate::helpers::{run_print, run_python_one};

#[test]
fn list_literal_basic() {
    assert_eq!(run_print("[1, 2, 3]"), "[1, 2, 3]");
}

#[test]
fn list_empty() {
    assert_eq!(run_print("[]"), "[]");
}

#[test]
fn list_index_first() {
    assert_eq!(run_print("[10, 20][0]"), "10");
}

#[test]
fn list_index_negative() {
    assert_eq!(run_print("[10, 20, 30][-1]"), "30");
}

#[test]
fn list_slice_step() {
    assert_eq!(run_print("[0, 1, 2, 3][::2]"), "[0, 2]");
}

#[test]
fn list_append_mutates() {
    assert_eq!(run_python_one("a = [1]\na.append(2)\nprint(a)\n"), "[1, 2]");
}

#[test]
fn list_extend() {
    assert_eq!(
        run_python_one("a = [1]\na.extend([2, 3])\nprint(a)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_insert_at_start() {
    assert_eq!(
        run_python_one("a = [2]\na.insert(0, 1)\nprint(a)\n"),
        "[1, 2]"
    );
}

#[test]
fn list_pop_last() {
    assert_eq!(run_python_one("a = [1, 2]\nprint(a.pop())\n"), "2");
}

#[test]
fn list_pop_index() {
    assert_eq!(run_python_one("a = [1, 2, 3]\nprint(a.pop(0))\n"), "1");
}

#[test]
fn list_remove_first_match() {
    assert_eq!(
        run_python_one("a = [1, 2, 1]\na.remove(1)\nprint(a)\n"),
        "[2, 1]"
    );
}

#[test]
fn list_clear() {
    assert_eq!(run_python_one("a = [1, 2]\na.clear()\nprint(a)\n"), "[]");
}

#[test]
fn list_copy_shallow() {
    assert_eq!(
        run_python_one("a = [1]\nb = a.copy()\nb.append(2)\nprint(a, b)\n"),
        "[1] [1, 2]"
    );
}

#[test]
fn list_concat() {
    assert_eq!(run_print("[1] + [2]"), "[1, 2]");
}

#[test]
fn list_repeat() {
    assert_eq!(run_print("[1] * 3"), "[1, 1, 1]");
}

#[test]
fn list_in_membership() {
    assert_eq!(run_print("2 in [1, 2, 3]"), "True");
}

#[test]
fn list_count() {
    assert_eq!(run_print("[1, 1, 2].count(1)"), "2");
}

#[test]
fn list_index_of_value() {
    assert_eq!(run_print("[10, 20, 30].index(20)"), "1");
}

#[test]
fn list_sort_ascending() {
    assert_eq!(
        run_python_one("a = [3, 1, 2]\na.sort()\nprint(a)\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_sort_descending() {
    assert_eq!(
        run_python_one("a = [3, 1, 2]\na.sort(reverse=True)\nprint(a)\n"),
        "[3, 2, 1]"
    );
}

#[test]
fn list_reverse_in_place() {
    assert_eq!(
        run_python_one("a = [1, 2]\na.reverse()\nprint(a)\n"),
        "[2, 1]"
    );
}

#[test]
fn list_reversed_builtin() {
    assert_eq!(run_print("list(reversed([1, 2, 3]))"), "[3, 2, 1]");
}

#[test]
fn list_comprehension_squares() {
    assert_eq!(run_print("[x*x for x in range(3)]"), "[0, 1, 4]");
}

#[test]
fn list_from_range() {
    assert_eq!(run_print("list(range(3))"), "[0, 1, 2]");
}

#[test]
fn list_from_tuple() {
    assert_eq!(run_print("list((1, 2))"), "[1, 2]");
}

#[test]
fn list_from_string() {
    assert_eq!(run_print("list('ab')"), "['a', 'b']");
}

#[test]
fn list_bool_nonempty() {
    assert_eq!(run_print("bool([0])"), "True");
}

#[test]
fn list_bool_empty() {
    assert_eq!(run_print("bool([])"), "False");
}

#[test]
fn list_len() {
    assert_eq!(run_print("len([1, 2, 3])"), "3");
}

#[test]
fn list_min_max() {
    assert_eq!(run_print("[min([3, 1, 2]), max([3, 1, 2])]"), "[1, 3]");
}

#[test]
fn list_sum() {
    assert_eq!(run_print("sum([1, 2, 3])"), "6");
}

#[test]
fn list_any_all() {
    assert_eq!(run_print("[any([0, 1]), all([1, 2])]"), "[True, True]");
}

#[test]
fn list_nested_access() {
    assert_eq!(run_print("[[1, 2], [3]][1][0]"), "3");
}

#[test]
fn list_unpack_three() {
    assert_eq!(run_python_one("a, b, c = [1, 2, 3]\nprint(b)\n"), "2");
}

#[test]
fn list_star_unpack() {
    assert_eq!(
        run_python_one("a, *mid, z = [1, 2, 3, 4]\nprint(a, mid, z)\n"),
        "1 [2, 3] 4"
    );
}

#[test]
fn list_equality() {
    assert_eq!(run_print("[1, 2] == [1, 2]"), "True");
}

#[test]
fn list_inequality() {
    assert_eq!(run_print("[1, 2] == [2, 1]"), "False");
}

#[test]
fn list_less_than_lexicographic() {
    assert_eq!(run_print("[1, 2] < [1, 3]"), "True");
}

#[test]
fn list_contains_sublist_false() {
    assert_eq!(run_print("[1, 2] in [[1, 2], [3]]"), "True");
}

#[test]
fn list_del_item() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\ndel a[1]\nprint(a)\n"),
        "[1, 3]"
    );
}

#[test]
fn list_slice_assignment() {
    assert_eq!(
        run_python_one("a = [1, 2, 3, 4]\na[1:3] = [9]\nprint(a)\n"),
        "[1, 9, 4]"
    );
}

#[test]
fn list_assign_to_slice_extend() {
    assert_eq!(
        run_python_one("a = [1, 2, 3]\na[1:1] = [9, 9]\nprint(a)\n"),
        "[1, 9, 9, 2, 3]"
    );
}

#[test]
fn list_filter_builtin() {
    assert_eq!(run_print("list(filter(None, [0, 1, 2]))"), "[1, 2]");
}

#[test]
fn list_map_builtin() {
    assert_eq!(run_print("list(map(str, [1, 2]))"), "['1', '2']");
}

#[test]
fn list_enumerate() {
    assert_eq!(run_print("list(enumerate(['a']))[0]"), "(0, 'a')");
}

#[test]
fn list_zip_two() {
    assert_eq!(
        run_print("list(zip([1, 2], ['a', 'b']))"),
        "[(1, 'a'), (2, 'b')]"
    );
}
