use crate::helpers::{run_print, run_python_one};

#[test]
fn enumerate_basic_index_value() {
    assert_eq!(
        run_print("list(enumerate(['a', 'b']))"),
        "[(0, 'a'), (1, 'b')]"
    );
}

#[test]
fn enumerate_with_start() {
    assert_eq!(
        run_print("list(enumerate(['a', 'b'], start=1))"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn enumerate_empty_iterable() {
    assert_eq!(run_print("list(enumerate([]))"), "[]");
}

#[test]
fn enumerate_string_chars() {
    assert_eq!(run_print("list(enumerate('ab'))"), "[(0, 'a'), (1, 'b')]");
}

#[test]
fn enumerate_range_values() {
    assert_eq!(
        run_print("list(enumerate(range(3)))"),
        "[(0, 0), (1, 1), (2, 2)]"
    );
}

#[test]
fn enumerate_for_loop_unpack() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i, v in enumerate(['x', 'y']):\n out.append(f'{i}:{v}')\nprint(out)\n"
        ),
        "['0:x', '1:y']"
    );
}

#[test]
fn enumerate_next_manual() {
    assert_eq!(
        run_python_one("it = enumerate(['a'])\nprint(next(it))\n"),
        "(0, 'a')"
    );
}

#[test]
fn zip_two_lists() {
    assert_eq!(
        run_print("list(zip([1, 2], ['a', 'b']))"),
        "[(1, 'a'), (2, 'b')]"
    );
}

#[test]
fn zip_three_lists() {
    assert_eq!(run_print("list(zip([1], [2], [3]))"), "[(1, 2, 3)]");
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
fn zip_for_loop_unpack() {
    assert_eq!(
        run_python_one("s = 0\nfor a, b in zip([1, 2], [10, 20]):\n s += a + b\nprint(s)\n"),
        "33"
    );
}

#[test]
fn zip_strict_equal_lengths() {
    assert_eq!(
        run_python_one(
            "try:\n list(zip([1, 2], [1], strict=True))\nexcept ValueError:\n print('val')\n"
        ),
        "val"
    );
}

#[test]
fn zip_longest_not_used_use_manual() {
    assert_eq!(
        run_python_one("pairs = list(zip([1, 2], ['a']))\nprint(len(pairs))\n"),
        "1"
    );
}

#[test]
fn enumerate_zip_combined() {
    assert_eq!(
        run_print("list(enumerate(zip([1, 2], ['a', 'b'])))"),
        "[(0, (1, 'a')), (1, (2, 'b'))]"
    );
}

#[test]
fn map_with_zip_unpack() {
    assert_eq!(
        run_print("list(map(lambda p: p[0] + len(p[1]), zip([1, 2], ['a', 'bb'])))"),
        "[2, 4]"
    );
}

#[test]
fn dict_from_zip() {
    assert_eq!(
        run_print("dict(zip(['a', 'b'], [1, 2]))"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn list_comp_with_enumerate_filter() {
    assert_eq!(
        run_print("[i for i, v in enumerate(['a', 'b', 'c']) if v == 'b']"),
        "[1]"
    );
}

#[test]
fn list_comp_with_zip() {
    assert_eq!(
        run_print("[a + b for a, b in zip([1, 2], [10, 20])]"),
        "[11, 22]"
    );
}

#[test]
fn enumerate_negative_start() {
    assert_eq!(run_print("list(enumerate(['x'], start=-1))"), "[(-1, 'x')]");
}

#[test]
fn enumerate_large_start() {
    assert_eq!(run_print("list(enumerate(['x'], start=100))[0][0]"), "100");
}

#[test]
fn zip_tuple_input() {
    assert_eq!(run_print("list(zip((1, 2), (3, 4)))"), "[(1, 3), (2, 4)]");
}

#[test]
fn zip_string_chars() {
    assert_eq!(
        run_print("list(zip('ab', 'xy'))"),
        "[('a', 'x'), ('b', 'y')]"
    );
}

#[test]
fn enumerate_dict_keys() {
    assert_eq!(
        run_print("list(enumerate({'a': 1, 'b': 2}))"),
        "[(0, 'a'), (1, 'b')]"
    );
}

#[test]
fn zip_dict_items_unpacked() {
    assert_eq!(
        run_print("list(zip({'a': 1, 'b': 2}.items()))"),
        "[(('a', 1),), (('b', 2),)]"
    );
}

#[test]
fn enumerate_slice_materialized() {
    assert_eq!(
        run_print("list(enumerate([10, 20, 30]))[1:]"),
        "[(1, 20), (2, 30)]"
    );
}

#[test]
fn zip_unzip_via_star() {
    assert_eq!(
        run_python_one(
            "pairs = [(1, 'a'), (2, 'b')]\nnums, letters = zip(*pairs)\nprint(list(nums), list(letters))\n"
        ),
        "[1, 2] ['a', 'b']"
    );
}

#[test]
fn enumerate_parallel_lists() {
    assert_eq!(
        run_python_one("names = ['a', 'b']\nscores = [1, 2]\nprint(list(zip(names, scores)))\n"),
        "[('a', 1), ('b', 2)]"
    );
}

#[test]
fn zip_with_range() {
    assert_eq!(
        run_print("list(zip(range(2), range(10, 12)))"),
        "[(0, 10), (1, 11)]"
    );
}

#[test]
fn enumerate_generator_expression() {
    assert_eq!(
        run_print("list(enumerate(x for x in range(2)))"),
        "[(0, 0), (1, 1)]"
    );
}

#[test]
fn zip_generator_expression() {
    assert_eq!(
        run_print("list(zip((x for x in range(2)), 'ab'))"),
        "[(0, 'a'), (1, 'b')]"
    );
}

#[test]
fn enumerate_bool_filter() {
    assert_eq!(
        run_print("[v for i, v in enumerate([0, 1, 2]) if v]"),
        "[1, 2]"
    );
}

#[test]
fn zip_sum_columns() {
    assert_eq!(
        run_python_one("cols = list(zip([1, 2], [3, 4]))\nprint(sum(a + b for a, b in cols))\n"),
        "10"
    );
}

#[test]
fn enumerate_find_index_of_match() {
    assert_eq!(
        run_python_one(
            "target = 'y'\nidx = next(i for i, v in enumerate(['x', 'y']) if v == target)\nprint(idx)\n"
        ),
        "1"
    );
}

#[test]
fn zip_unequal_length_default_behavior() {
    assert_eq!(run_print("len(list(zip(range(5), range(2))))"), "2");
}

#[test]
fn enumerate_on_set_sorted_keys() {
    assert_eq!(
        run_print("list(enumerate(sorted({3, 1, 2})))"),
        "[(0, 1), (1, 2), (2, 3)]"
    );
}

#[test]
fn zip_with_single_element_lists() {
    assert_eq!(run_print("list(zip([1], ['a']))"), "[(1, 'a')]");
}

#[test]
fn enumerate_count_filtered_items() {
    assert_eq!(
        run_python_one("print(sum(1 for i, v in enumerate([1, 2, 3]) if v > 1))\n"),
        "2"
    );
}

#[test]
fn zip_build_lookup_dict() {
    assert_eq!(
        run_print("dict(zip(['k1', 'k2'], [9, 8]))"),
        "{'k1': 9, 'k2': 8}"
    );
}

#[test]
fn enumerate_string_find_position() {
    assert_eq!(
        run_python_one("s = 'banana'\nprint(next(i for i, ch in enumerate(s) if ch == 'n'))\n"),
        "2"
    );
}

#[test]
fn zip_transpose_rows() {
    assert_eq!(run_print("list(zip([1, 2], [3, 4]))"), "[(1, 3), (2, 4)]");
}

#[test]
fn enumerate_list_comp_strings() {
    assert_eq!(
        run_print("[f'{i}{v}' for i, v in enumerate(['a', 'b'])]"),
        "['0a', '1b']"
    );
}

#[test]
fn zip_filter_pairs() {
    assert_eq!(
        run_print("[(a, b) for a, b in zip([1, 2, 3], [0, 2, 0]) if b]"),
        "[(2, 2)]"
    );
}

#[test]
fn enumerate_reverse_not_builtin_use_slice() {
    assert_eq!(
        run_print("list(enumerate([1, 2, 3]))[::-1]"),
        "[(2, 3), (1, 2), (0, 1)]"
    );
}

#[test]
fn zip_chain_two_zips() {
    assert_eq!(
        run_print("list(zip([1, 2], 'ab')) + list(zip([3], 'c'))"),
        "[(1, 'a'), (2, 'b'), (3, 'c')]"
    );
}
