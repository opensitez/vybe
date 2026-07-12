use crate::helpers::{run_print, run_python_one};

#[test]
fn list_comp_squares_basic() {
    assert_eq!(run_print("[x * x for x in range(4)]"), "[0, 1, 4, 9]");
}

#[test]
fn list_comp_filters_evens() {
    assert_eq!(
        run_print("[x for x in range(6) if x % 2 == 0]"),
        "[0, 2, 4]"
    );
}

#[test]
fn list_comp_nested_flattens_pairs() {
    assert_eq!(
        run_print("[b for a in [1, 2] for b in [a, a + 10]]"),
        "[1, 11, 2, 12]"
    );
}

#[test]
fn list_comp_if_at_end_filters() {
    assert_eq!(run_print("[x for x in range(5) if x > 2]"), "[3, 4]");
}

#[test]
fn list_comp_transform_strings() {
    assert_eq!(run_print("[s.upper() for s in ['a', 'b']]"), "['A', 'B']");
}

#[test]
fn list_comp_with_condition_on_transform() {
    assert_eq!(
        run_print("[x * 2 for x in range(5) if x % 2 == 1]"),
        "[2, 6]"
    );
}

#[test]
fn list_comp_from_string_chars() {
    assert_eq!(run_print("[c for c in 'hi']"), "['h', 'i']");
}

#[test]
fn list_comp_dict_keys_list() {
    assert_eq!(run_print("list(k for k in {'a': 1, 'b': 2})"), "['a', 'b']");
}

#[test]
fn list_comp_enumerate_style() {
    assert_eq!(
        run_python_one("print([i for i, v in enumerate(['x', 'y']) if v == 'y'][0])\n"),
        "1"
    );
}

#[test]
fn list_comp_zip_pairs() {
    assert_eq!(
        run_print("[a + b for a, b in zip([1, 2], [10, 20])]"),
        "[11, 22]"
    );
}

#[test]
fn list_comp_nested_matrix_rows() {
    assert_eq!(
        run_print("[[j for j in range(2)] for i in range(2)]"),
        "[[0, 1], [0, 1]]"
    );
}

#[test]
fn list_comp_filter_none_values() {
    assert_eq!(
        run_print("[x for x in [0, 1, None, 2] if x is not None]"),
        "[0, 1, 2]"
    );
}

#[test]
fn list_comp_truthy_strings() {
    assert_eq!(run_print("[s for s in ['', 'a', ''] if s]"), "['a']");
}

#[test]
fn list_comp_length_filter() {
    assert_eq!(
        run_print("[w for w in ['hi', 'hey', 'yo'] if len(w) == 2]"),
        "['hi', 'yo']"
    );
}

#[test]
fn list_comp_duplicate_range_values() {
    assert_eq!(run_print("[x for x in [1, 1, 2, 2, 3]]"), "[1, 1, 2, 2, 3]");
}

#[test]
fn list_comp_unique_with_set_cast() {
    assert_eq!(run_print("sorted({x for x in [3, 1, 2, 1]})"), "[1, 2, 3]");
}

#[test]
fn list_comp_conditional_expression_inside() {
    assert_eq!(
        run_print("[('even' if x % 2 == 0 else 'odd') for x in range(3)]"),
        "['even', 'odd', 'even']"
    );
}

#[test]
fn list_comp_slice_of_range() {
    assert_eq!(run_print("[x for x in range(10) if x < 3]"), "[0, 1, 2]");
}

#[test]
fn list_comp_negative_indices_via_enumerate() {
    assert_eq!(
        run_print("[i for i, v in enumerate([10, 20, 30]) if v == 30]"),
        "[2]"
    );
}

#[test]
fn list_comp_from_split_words() {
    assert_eq!(
        run_print("[w for w in 'a,b,c'.split(',')]"),
        "['a', 'b', 'c']"
    );
}

#[test]
fn list_comp_map_like_doubling() {
    assert_eq!(run_print("[n * 2 for n in [1, 2, 3]]"), "[2, 4, 6]");
}

#[test]
fn list_comp_filter_map_combined() {
    assert_eq!(
        run_print("[n * n for n in range(6) if n % 2 == 0]"),
        "[0, 4, 16]"
    );
}

#[test]
fn list_comp_nested_filter() {
    assert_eq!(
        run_print("[x for x in [1, 2, 3, 4] if x > 1 if x < 4]"),
        "[2, 3]"
    );
}

#[test]
fn list_comp_on_tuple_source() {
    assert_eq!(run_print("[x for x in (1, 2, 3)]"), "[1, 2, 3]");
}

#[test]
fn list_comp_on_set_source_sorted() {
    assert_eq!(run_print("sorted([x for x in {3, 1, 2}])"), "[1, 2, 3]");
}

#[test]
fn list_comp_string_digits_only() {
    assert_eq!(
        run_print("[c for c in 'a1b2' if c.isdigit()]"),
        "['1', '2']"
    );
}

#[test]
fn list_comp_build_lookup_table() {
    assert_eq!(run_print("[i * 10 for i in range(3)]"), "[0, 10, 20]");
}

#[test]
fn list_comp_empty_when_filter_excludes_all() {
    assert_eq!(run_print("[x for x in range(3) if x > 10]"), "[]");
}

#[test]
fn list_comp_single_element() {
    assert_eq!(run_print("[42]"), "[42]");
}

#[test]
fn list_comp_from_empty_range() {
    assert_eq!(run_print("[x for x in range(0)]"), "[]");
}

#[test]
fn list_comp_replicate_constant() {
    assert_eq!(run_print("[0 for _ in range(4)]"), "[0, 0, 0, 0]");
}

#[test]
fn list_comp_indexed_comprehension() {
    assert_eq!(
        run_print("[i for i in range(5) if i % 2 == 0]"),
        "[0, 2, 4]"
    );
}

#[test]
fn list_comp_pairs_from_two_lists() {
    assert_eq!(
        run_print("[(a, b) for a in [1] for b in [2, 3]]"),
        "[(1, 2), (1, 3)]"
    );
}

#[test]
fn list_comp_filter_positive_numbers() {
    assert_eq!(run_print("[x for x in [-1, 0, 2] if x > 0]"), "[2]");
}

#[test]
fn list_comp_string_length_map() {
    assert_eq!(run_print("[len(w) for w in ['a', 'ab']]"), "[1, 2]");
}

#[test]
fn list_comp_bool_coercion_filter() {
    assert_eq!(run_print("[x for x in [0, 1, 2] if x]"), "[1, 2]");
}

#[test]
fn list_comp_chars_upper_filtered() {
    assert_eq!(run_print("[c.upper() for c in 'ab' if c != 'b']"), "['A']");
}

#[test]
fn list_comp_modulo_classes() {
    assert_eq!(run_print("[x for x in range(6) if x % 3 == 0]"), "[0, 3]");
}

#[test]
fn list_comp_reversed_range_via_slice() {
    assert_eq!(run_print("[x for x in range(3)][::-1]"), "[2, 1, 0]");
}

#[test]
fn list_comp_join_strings_after() {
    assert_eq!(
        run_python_one("print('-'.join([str(x) for x in range(3)]))\n"),
        "0-1-2"
    );
}

#[test]
fn list_comp_sum_of_squares() {
    assert_eq!(
        run_python_one("print(sum([x * x for x in range(4)]))\n"),
        "14"
    );
}

#[test]
fn list_comp_any_match() {
    assert_eq!(
        run_python_one("print(any([x > 2 for x in [1, 2, 3]]))\n"),
        "True"
    );
}

#[test]
fn list_comp_all_match() {
    assert_eq!(
        run_python_one("print(all([x > 0 for x in [1, 2, 3]]))\n"),
        "True"
    );
}

#[test]
fn list_comp_max_of_transformed() {
    assert_eq!(
        run_python_one("print(max([len(s) for s in ['a', 'bbb']]))\n"),
        "3"
    );
}

#[test]
fn list_comp_min_of_transformed() {
    assert_eq!(
        run_python_one("print(min([len(s) for s in ['aa', 'b']]))\n"),
        "1"
    );
}
