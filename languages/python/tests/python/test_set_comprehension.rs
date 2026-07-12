use crate::helpers::{run_print, run_python_one};

#[test]
fn set_comp_squares() {
    assert_eq!(
        run_print("sorted({x * x for x in range(4)})"),
        "[0, 1, 4, 9]"
    );
}

#[test]
fn set_comp_filtered_evens() {
    assert_eq!(
        run_print("sorted({x for x in range(6) if x % 2 == 0})"),
        "[0, 2, 4]"
    );
}

#[test]
fn set_comp_from_string_unique_chars() {
    assert_eq!(
        run_print("sorted({c for c in 'banana'})"),
        "['a', 'b', 'n']"
    );
}

#[test]
fn set_comp_exclude_spaces() {
    assert_eq!(
        run_print("sorted({c for c in 'a b' if c != ' '})"),
        "['a', 'b']"
    );
}

#[test]
fn set_comp_nested_loops() {
    assert_eq!(
        run_print("sorted({a + b for a in [1, 2] for b in [10, 20]})"),
        "[11, 12, 21, 22]"
    );
}

#[test]
fn set_comp_from_list_removes_dupes() {
    assert_eq!(
        run_print("sorted({x for x in [1, 1, 2, 3, 2]})"),
        "[1, 2, 3]"
    );
}

#[test]
fn set_comp_modulo_classes() {
    assert_eq!(run_print("sorted({x % 3 for x in range(6)})"), "[0, 1, 2]");
}

#[test]
fn set_comp_len_of_strings() {
    assert_eq!(
        run_print("sorted({len(s) for s in ['a', 'bb', 'a']})"),
        "[1, 2]"
    );
}

#[test]
fn set_comp_uppercase_words() {
    assert_eq!(
        run_print("sorted({w.upper() for w in ['a', 'b']})"),
        "['A', 'B']"
    );
}

#[test]
fn set_comp_from_dict_keys() {
    assert_eq!(
        run_print("sorted({k for k in {'x': 1, 'y': 2}})"),
        "['x', 'y']"
    );
}

#[test]
fn set_comp_from_dict_values() {
    assert_eq!(
        run_print("sorted({v for v in {'a': 1, 'b': 2}.values()})"),
        "[1, 2]"
    );
}

#[test]
fn set_comp_zip_pairs_as_tuples() {
    assert_eq!(
        run_print("sorted({(a, b) for a, b in zip([1, 2], [3, 4])})"),
        "[(1, 3), (2, 4)]"
    );
}

#[test]
fn set_comp_truthy_values() {
    assert_eq!(
        run_print("sorted({x for x in [0, 1, 2, 0] if x})"),
        "[1, 2]"
    );
}

#[test]
fn set_comp_negative_numbers() {
    assert_eq!(
        run_print("sorted({x for x in [-1, -2, 1] if x < 0})"),
        "[-2, -1]"
    );
}

#[test]
fn set_comp_float_rounded_int_cast() {
    assert_eq!(
        run_print("sorted({int(x) for x in [1.1, 1.9, 2.1]})"),
        "[1, 2]"
    );
}

#[test]
fn set_comp_powers_of_two() {
    assert_eq!(
        run_print("sorted({2 ** n for n in range(4)})"),
        "[1, 2, 4, 8]"
    );
}

#[test]
fn set_comp_division_results() {
    assert_eq!(run_print("sorted({x // 2 for x in range(5)})"), "[0, 1, 2]");
}

#[test]
fn set_comp_chars_isalpha() {
    assert_eq!(
        run_print("sorted({c for c in 'a1b2' if c.isalpha()})"),
        "['a', 'b']"
    );
}

#[test]
fn set_comp_chars_isdigit() {
    assert_eq!(
        run_print("sorted({c for c in 'a1b2' if c.isdigit()})"),
        "['1', '2']"
    );
}

#[test]
fn set_comp_from_split_words() {
    assert_eq!(
        run_print("sorted({w for w in 'hi,ho'.split(',')})"),
        "['hi', 'ho']"
    );
}

#[test]
fn set_comp_range_step() {
    assert_eq!(
        run_print("sorted({x for x in range(0, 10, 3)})"),
        "[0, 3, 6, 9]"
    );
}

#[test]
fn set_comp_conditional_expression() {
    assert_eq!(
        run_print("sorted({('even' if x % 2 == 0 else 'odd') for x in range(3)})"),
        "['even', 'odd']"
    );
}

#[test]
fn set_comp_empty_when_filter_false() {
    assert_eq!(run_print("{x for x in range(3) if x > 10}"), "set()");
}

#[test]
fn set_comp_singleton() {
    assert_eq!(run_print("{42}"), "{42}");
}

#[test]
fn set_comp_union_with_literal() {
    assert_eq!(
        run_print("sorted({x for x in range(2)} | {2, 3})"),
        "[0, 1, 2, 3]"
    );
}

#[test]
fn set_comp_intersection_with_range() {
    assert_eq!(
        run_print("sorted({x for x in range(5)} & {2, 3, 9})"),
        "[2, 3]"
    );
}

#[test]
fn set_comp_difference() {
    assert_eq!(
        run_print("sorted({x for x in range(5)} - {0, 1})"),
        "[2, 3, 4]"
    );
}

#[test]
fn set_comp_symmetric_difference() {
    assert_eq!(
        run_print("sorted({x for x in range(3)} ^ {1, 2, 4})"),
        "[0, 4]"
    );
}

#[test]
fn set_comp_frozenset_elements() {
    assert_eq!(
        run_print("sorted({frozenset({i}) for i in range(2)})"),
        "[frozenset({0}), frozenset({1})]"
    );
}

#[test]
fn set_comp_tuple_elements_hashable() {
    assert_eq!(
        run_print("sorted({(i, i + 1) for i in range(2)})"),
        "[(0, 1), (1, 2)]"
    );
}

#[test]
fn set_comp_abs_values() {
    assert_eq!(run_print("sorted({abs(x) for x in [-2, -1, 1]})"), "[1, 2]");
}

#[test]
fn set_comp_string_stripped() {
    assert_eq!(
        run_print("sorted({s.strip() for s in [' a', 'b ']})"),
        "['a', 'b']"
    );
}

#[test]
fn set_comp_enumerate_indices() {
    assert_eq!(
        run_print("sorted({i for i, _ in enumerate(['a', 'b', 'a'])})"),
        "[0, 1, 2]"
    );
}

#[test]
fn set_comp_enumerate_values() {
    assert_eq!(
        run_print("sorted({v for _, v in enumerate(['x', 'y'])})"),
        "['x', 'y']"
    );
}

#[test]
fn set_comp_map_result() {
    assert_eq!(
        run_print("sorted({x * 2 for x in map(int, ['1', '2', '2'])})"),
        "[2, 4]"
    );
}

#[test]
fn set_comp_filter_on_range() {
    assert_eq!(
        run_print("sorted({x for x in filter(lambda n: n % 2 == 1, range(6))})"),
        "[1, 3, 5]"
    );
}

#[test]
fn set_comp_any_length_predicate() {
    assert_eq!(
        run_python_one("s = {len(w) for w in ['a', 'bb', 'ccc']}\nprint(len(s))\n"),
        "3"
    );
}

#[test]
fn set_comp_all_positive() {
    assert_eq!(
        run_python_one("s = {x for x in range(1, 4)}\nprint(all(v > 0 for v in s))\n"),
        "True"
    );
}

#[test]
fn set_comp_max_element() {
    assert_eq!(run_python_one("print(max({x for x in range(5)}))\n"), "4");
}

#[test]
fn set_comp_min_element() {
    assert_eq!(run_python_one("print(min({x for x in range(5)}))\n"), "0");
}

#[test]
fn set_comp_sum_elements() {
    assert_eq!(run_python_one("print(sum({x for x in range(4)}))\n"), "6");
}

#[test]
fn set_comp_in_list_membership() {
    assert_eq!(
        run_python_one("s = {x for x in [1, 2]}\nprint(2 in s)\n"),
        "True"
    );
}

#[test]
fn set_comp_subset_check() {
    assert_eq!(
        run_python_one("a = {x for x in range(2)}\nb = {x for x in range(3)}\nprint(a <= b)\n"),
        "True"
    );
}

#[test]
fn set_comp_proper_superset() {
    assert_eq!(
        run_python_one("a = {x for x in range(3)}\nb = {x for x in range(2)}\nprint(a > b)\n"),
        "True"
    );
}

#[test]
fn set_comp_disjoint() {
    assert_eq!(
        run_python_one(
            "a = {x for x in range(2)}\nb = {x for x in range(2, 4)}\nprint(a.isdisjoint(b))\n"
        ),
        "True"
    );
}
