use crate::helpers::{run_python_one, run_print};

#[test]
fn dict_comp_range_doubled_values() {
    assert_eq!(run_print("{k: k * 2 for k in range(3)}"), "{0: 0, 1: 2, 2: 4}");
}

#[test]
fn dict_comp_string_lengths() {
    assert_eq!(
        run_print("{s: len(s) for s in ['a', 'ab']}"),
        "{'a': 1, 'ab': 2}"
    );
}

#[test]
fn dict_comp_filtered_positive() {
    assert_eq!(
        run_print("{x: x for x in range(5) if x > 2}"),
        "{3: 3, 4: 4}"
    );
}

#[test]
fn dict_comp_from_pairs_list() {
    assert_eq!(
        run_print("{k: v for k, v in [('a', 1), ('b', 2)]}"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn dict_comp_zip_two_lists() {
    assert_eq!(
        run_print("{a: b for a, b in zip(['x', 'y'], [1, 2])}"),
        "{'x': 1, 'y': 2}"
    );
}

#[test]
fn dict_comp_enumerate_indices() {
    assert_eq!(
        run_print("{i: v for i, v in enumerate(['a', 'b'])}"),
        "{0: 'a', 1: 'b'}"
    );
}

#[test]
fn dict_comp_transform_values() {
    assert_eq!(
        run_print("{k: k.upper() for k in ['a', 'b']}"),
        "{'a': 'A', 'b': 'B'}"
    );
}

#[test]
fn dict_comp_keys_from_string_chars() {
    assert_eq!(
        run_print("{c: ord(c) for c in 'ab' if c == 'a'}"),
        "{'a': 97}"
    );
}

#[test]
fn dict_comp_nested_loop() {
    assert_eq!(
        run_print("{f'{a}{b}': a + b for a in [1] for b in [2, 3]}"),
        "{'12': 3, '13': 4}"
    );
}

#[test]
fn dict_comp_conditional_value() {
    assert_eq!(
        run_print("{x: ('even' if x % 2 == 0 else 'odd') for x in range(3)}"),
        "{0: 'even', 1: 'odd', 2: 'even'}"
    );
}

#[test]
fn dict_comp_empty_filter() {
    assert_eq!(run_print("{x: x for x in range(3) if x > 10}"), "{}");
}

#[test]
fn dict_comp_from_set_keys() {
    assert_eq!(
        run_print("{k: k * k for k in {1, 2}}"),
        "{1: 1, 2: 4}"
    );
}

#[test]
fn dict_comp_bool_keys() {
    assert_eq!(
        run_print("{str(v): v for v in [True, False]}"),
        "{'True': True, 'False': False}"
    );
}

#[test]
fn dict_comp_value_list_comp() {
    assert_eq!(
        run_print("{k: [k] for k in range(2)}"),
        "{0: [0], 1: [1]}"
    );
}

#[test]
fn dict_comp_merge_style_update() {
    assert_eq!(
        run_python_one("base = {'a': 1}\nbase.update({k: k for k in range(2)})\nprint(base)\n"),
        "{'a': 1, 0: 0, 1: 1}"
    );
}

#[test]
fn dict_comp_lookup_then_invert() {
    assert_eq!(
        run_print("{v: k for k, v in {'a': 1, 'b': 2}.items()}"),
        "{1: 'a', 2: 'b'}"
    );
}

#[test]
fn dict_comp_sum_values() {
    assert_eq!(
        run_python_one("d = {i: i for i in range(4)}\nprint(sum(d.values()))\n"),
        "6"
    );
}

#[test]
fn dict_comp_keys_sorted() {
    assert_eq!(
        run_python_one("d = {x: x for x in [3, 1, 2]}\nprint(sorted(d.keys()))\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn dict_comp_duplicate_keys_last_wins() {
    assert_eq!(
        run_print("{1: 'a', 1: 'b'}"),
        "{1: 'b'}"
    );
}

#[test]
fn dict_comp_from_tuple_keys() {
    assert_eq!(
        run_print("{(i,): i for i in range(2)}"),
        "{(0,): 0, (1,): 1}"
    );
}

#[test]
fn dict_comp_filter_map_items() {
    assert_eq!(
        run_print("{k: v for k, v in {'a': 1, 'b': 0}.items() if v}"),
        "{'a': 1}"
    );
}

#[test]
fn dict_comp_string_digit_map() {
    assert_eq!(
        run_print("{c: int(c) for c in '12'}"),
        "{'1': 1, '2': 2}"
    );
}

#[test]
fn dict_comp_abs_values() {
    assert_eq!(
        run_print("{x: abs(x) for x in [-1, 2]}"),
        "{-1: 1, 2: 2}"
    );
}

#[test]
fn dict_comp_len_on_values() {
    assert_eq!(
        run_python_one("d = {k: len(k) for k in ['a', 'bbb']}\nprint(d['bbb'])\n"),
        "3"
    );
}

#[test]
fn dict_comp_nested_dict_value() {
    assert_eq!(
        run_print("{k: {'v': k} for k in range(2)}"),
        "{0: {'v': 0}, 1: {'v': 1}}"
    );
}

#[test]
fn dict_comp_identity_on_strings() {
    assert_eq!(
        run_print("{w: w for w in ['hi', 'yo']}"),
        "{'hi': 'hi', 'yo': 'yo'}"
    );
}

#[test]
fn dict_comp_modulo_classes() {
    assert_eq!(
        run_print("{x: x % 3 for x in range(6)}"),
        "{0: 0, 1: 1, 2: 2, 3: 0, 4: 1, 5: 2}"
    );
}

#[test]
fn dict_comp_power_table() {
    assert_eq!(
        run_print("{n: n ** 2 for n in range(4)}"),
        "{0: 0, 1: 1, 2: 4, 3: 9}"
    );
}

#[test]
fn dict_comp_join_key_parts() {
    assert_eq!(
        run_print("{'-'.join([str(a), str(b)]): a + b for a, b in [(1, 2)]}"),
        "{'1-2': 3}"
    );
}

#[test]
fn dict_comp_any_value_truthy() {
    assert_eq!(
        run_python_one("d = {i: bool(i) for i in range(3)}\nprint(any(d.values()))\n"),
        "True"
    );
}

#[test]
fn dict_comp_all_keys_strings() {
    assert_eq!(
        run_python_one("d = {str(i): i for i in range(2)}\nprint(all(isinstance(k, str) for k in d))\n"),
        "True"
    );
}

#[test]
fn dict_comp_from_split() {
    assert_eq!(
        run_print("{p: len(p) for p in 'a,b'.split(',')}"),
        "{'a': 1, 'b': 1}"
    );
}

#[test]
fn dict_comp_float_keys_cast_str() {
    assert_eq!(
        run_print("{str(x): x for x in [1.5]}"),
        "{'1.5': 1.5}"
    );
}

#[test]
fn dict_comp_list_value_mutation_independent() {
    assert_eq!(
        run_python_one("d = {i: [] for i in range(2)}\nd[0].append(1)\nprint(d[1])\n"),
        "[]"
    );
}

#[test]
fn dict_comp_filter_none_values_out() {
    assert_eq!(
        run_print("{i: v for i, v in enumerate([0, None, 2]) if v is not None}"),
        "{0: 0, 2: 2}"
    );
}

#[test]
fn dict_comp_chars_to_bool() {
    assert_eq!(
        run_print("{c: c.isalpha() for c in 'a1'}"),
        "{'a': True, '1': False}"
    );
}

#[test]
fn dict_comp_range_step() {
    assert_eq!(
        run_print("{x: x * 10 for x in range(0, 10, 5)}"),
        "{0: 0, 5: 50}"
    );
}

#[test]
fn dict_comp_max_key_lookup() {
    assert_eq!(
        run_python_one("d = {i: i * i for i in range(4)}\nprint(d[max(d)])\n"),
        "9"
    );
}

#[test]
fn dict_comp_get_with_default() {
    assert_eq!(
        run_python_one("d = {k: k for k in range(2)}\nprint(d.get(9, 'missing'))\n"),
        "missing"
    );
}

#[test]
fn dict_comp_items_roundtrip() {
    assert_eq!(
        run_python_one("d = {k: v * 2 for k, v in {'a': 1}.items()}\nprint(list(d.items())[0])\n"),
        "('a', 2)"
    );
}

#[test]
fn dict_comp_nested_conditional_keys() {
    assert_eq!(
        run_print("{('pos' if x > 0 else 'neg'): x for x in [-1, 1]}"),
        "{'neg': -1, 'pos': 1}"
    );
}

#[test]
fn dict_comp_sorted_values_list() {
    assert_eq!(
        run_python_one("d = {x: -x for x in [3, 1, 2]}\nprint(sorted(d.values()))\n"),
        "[-3, -2, -1]"
    );
}

#[test]
fn dict_comp_hashed_tuple_key_count() {
    assert_eq!(
        run_python_one("d = {(a, b): a + b for a in [1] for b in [2, 3]}\nprint(len(d))\n"),
        "2"
    );
}

#[test]
fn dict_comp_string_upper_keys() {
    assert_eq!(
        run_print("{k.upper(): v for k, v in {'a': 1}.items()}"),
        "{'A': 1}"
    );
}

#[test]
fn dict_comp_value_is_length_of_key() {
    assert_eq!(
        run_print("{word: len(word) for word in ['go', 'stop']}"),
        "{'go': 2, 'stop': 4}"
    );
}
