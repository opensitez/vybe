use crate::helpers::{run_print, run_python_one};

#[test]
fn zip_two_lists_pairs() {
    assert_eq!(
        run_print("list(zip([1, 2], ['a', 'b']))"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn zip_three_lists() {
    assert_eq!(
        run_print("list(zip([1, 2], ['a', 'b'], [True, False]))"),
        "[(1, 'a', True), (2, 'b', False)]"
    );
}

#[test]
fn zip_stops_at_shortest() {
    assert_eq!(run_print("list(zip([1, 2, 3], ['a']))"), "[(1, 'a')]");
}

#[test]
fn zip_empty_list() {
    assert_eq!(run_print("list(zip([], [1]))"), "[]");
}

#[test]
fn zip_with_range() {
    assert_eq!(
        run_print("list(zip(range(3), 'abc'))"),
        "[(0, 'a'), (1, 'b'), (2, 'c')]"
    );
}

#[test]
fn zip_unpack_in_for_loop() {
    assert_eq!(
        run_python_one("s = 0\nfor a, b in zip([1, 2], [10, 20]):\n s += a + b\nprint(s)\n"),
        "33"
    );
}

#[test]
fn zip_to_dict_constructor() {
    assert_eq!(
        run_print("dict(zip(['a', 'b'], [1, 2]))"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn zip_longest_fillvalue() {
    assert_eq!(
        run_print("list(__import__('itertools').zip_longest([1, 2], [3], fillvalue=0))"),
        "[(1, 3), (2, 0)]"
    );
}

#[test]
fn zip_star_unzip() {
    assert_eq!(
        run_print("list(zip(*[(1, 'a'), (2, 'b')]))"),
        "[(1, 2), ('a', 'b')]"
    );
}

#[test]
fn zip_with_enumerate_combo() {
    assert_eq!(
        run_print("list(zip(enumerate('ab'), [10, 20]))"),
        "[((0, 'a'), 10), ((1, 'b'), 20)]"
    );
}

#[test]
fn zip_list_comprehension_sum_products() {
    assert_eq!(
        run_print("sum(a * b for a, b in zip([1, 2, 3], [4, 5, 6]))"),
        "32"
    );
}

#[test]
fn zip_strings_parallel_chars() {
    assert_eq!(
        run_print("list(zip('ab', 'xy'))"),
        "[('a', 'x'), ('b', 'y')]"
    );
}

#[test]
fn zip_single_iterable_returns_one_tuples() {
    assert_eq!(run_print("list(zip([1, 2, 3]))"), "[(1,), (2,), (3,)]");
}

#[test]
fn zip_with_tuple_inputs() {
    assert_eq!(run_print("list(zip((1, 2), (3, 4)))"), "[(1, 3), (2, 4)]");
}

#[test]
fn zip_filter_pairs_by_condition() {
    assert_eq!(
        run_print("[(a, b) for a, b in zip([1, 2, 3], [3, 2, 1]) if a != b]"),
        "[(1, 3), (3, 1)]"
    );
}

#[test]
fn zip_dict_keys_and_values() {
    assert_eq!(
        run_python_one("d = {'x': 9, 'y': 8}\nprint(list(zip(d, d.values())))\n"),
        "[('x', 9), ('y', 8)]"
    );
}

#[test]
fn zip_parallel_increment_lists() {
    assert_eq!(
        run_print("list(zip([0, 1, 2], [1, 2, 3], [2, 3, 4]))"),
        "[(0, 1, 2), (1, 2, 3), (2, 3, 4)]"
    );
}

#[test]
fn zip_empty_zip_with_nonempty_gives_empty() {
    assert_eq!(run_print("list(zip([], range(5)))"), "[]");
}

#[test]
fn zip_bool_and_int_pairs() {
    assert_eq!(
        run_print("list(zip([True, False], [1, 0]))"),
        "[(True, 1), (False, 0)]"
    );
}

#[test]
fn zip_nested_list_elements() {
    assert_eq!(
        run_print("list(zip([[1], [2]], [[3], [4]]))"),
        "[([1], [3]), ([2], [4])]"
    );
}

#[test]
fn zip_map_transpose_rows() {
    assert_eq!(
        run_print("list(zip([1, 2, 3], [4, 5, 6]))"),
        "[(1, 4), (2, 5), (3, 6)]"
    );
}

#[test]
fn zip_with_strict_length_mismatch_raises() {
    assert_eq!(
        run_python_one(
            "try:\n list(zip([1, 2], [1], strict=True))\n print('ok')\nexcept ValueError:\n print('ValueError')\n"
        ),
        "ValueError"
    );
}

#[test]
fn zip_strict_equal_lengths_ok() {
    assert_eq!(
        run_print("list(zip([1, 2], [3, 4], strict=True))"),
        "[(1, 3), (2, 4)]"
    );
}

#[test]
fn zip_for_matrix_row_col_indices() {
    assert_eq!(
        run_python_one("pairs = list(zip(range(2), range(10, 12)))\nprint(pairs[1])\n"),
        "(1, 11)"
    );
}

#[test]
fn zip_generator_consumed_once() {
    assert_eq!(
        run_python_one("it = zip([1], [2])\nprint(list(it))\n"),
        "[(1, 2)]"
    );
}

#[test]
fn zip_with_set_and_list_same_length() {
    assert_eq!(
        run_python_one("print(len(list(zip({1, 2}, [3, 4]))))\n"),
        "2"
    );
}

#[test]
fn zip_align_names_and_scores() {
    assert_eq!(
        run_python_one(
            "names = ['Ann', 'Bob']\nscores = [90, 85]\nprint(dict(zip(names, scores))['Bob'])\n"
        ),
        "85"
    );
}

#[test]
fn zip_longest_without_fillvalue_pads_none() {
    assert_eq!(
        run_print("list(__import__('itertools').zip_longest([1], [2, 3]))"),
        "[(1, 2), (None, 3)]"
    );
}

#[test]
fn zip_parallel_string_number_formatting() {
    assert_eq!(
        run_print("[f'{a}:{b}' for a, b in zip('ab', [1, 2])]"),
        "['a:1', 'b:2']"
    );
}

#[test]
fn zip_reversed_inputs() {
    assert_eq!(
        run_print("list(zip(reversed([1, 2]), reversed([3, 4])))"),
        "[(2, 4), (1, 3)]"
    );
}

#[test]
fn zip_with_slice_inputs() {
    assert_eq!(
        run_print("list(zip([1, 2, 3][:2], 'xy'))"),
        "[(1, 'x'), (2, 'y')]"
    );
}

#[test]
fn zip_chain_manual_unzip_first_column() {
    assert_eq!(
        run_python_one("rows = [(1, 'a'), (2, 'b')]\ncols = list(zip(*rows))\nprint(cols[0])\n"),
        "(1, 2)"
    );
}

#[test]
fn zip_empty_strict_ok() {
    assert_eq!(run_print("list(zip([], [], strict=True))"), "[]");
}

#[test]
fn zip_float_and_int_pairs() {
    assert_eq!(
        run_print("list(zip([1.5, 2.5], [1, 2]))"),
        "[(1.5, 1), (2.5, 2)]"
    );
}

#[test]
fn zip_in_dict_comprehension() {
    assert_eq!(
        run_print("{a: b for a, b in zip(['x', 'y'], [1, 2])}"),
        "{'x': 1, 'y': 2}"
    );
}

#[test]
fn zip_three_way_all_empty() {
    assert_eq!(run_print("list(zip([], [], []))"), "[]");
}

#[test]
fn zip_with_bytes_and_ints_length_mismatch() {
    assert_eq!(
        run_print("list(zip(b'ab', [1, 2, 3]))"),
        "[(97, 1), (98, 2)]"
    );
}

#[test]
fn zip_sorted_unique_keys_with_values() {
    assert_eq!(
        run_python_one(
            "d = {'b': 2, 'a': 1}\nprint(list(zip(sorted(d), [d[k] for k in sorted(d)])))\n"
        ),
        "[('a', 1), ('b', 2)]"
    );
}

#[test]
fn zip_parallel_accumulate_running_total() {
    assert_eq!(
        run_python_one(
            "total = 0\nfor v, delta in zip([1, 2, 3], [1, 1, 1]):\n total += v * delta\nprint(total)\n"
        ),
        "6"
    );
}
