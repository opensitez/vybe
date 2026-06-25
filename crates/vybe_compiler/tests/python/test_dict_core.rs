use crate::helpers::{run_python_one, run_print};

#[test]
fn dict_literal_two_keys() {
    assert_eq!(run_print("{'a': 1, 'b': 2}"), "{'a': 1, 'b': 2}");
}

#[test]
fn dict_empty() {
    assert_eq!(run_print("{}"), "{}");
}

#[test]
fn dict_get_existing() {
    assert_eq!(run_print("{'x': 9}.get('x')"), "9");
}

#[test]
fn dict_get_missing_default() {
    assert_eq!(run_print("{}.get('k', 0)"), "0");
}

#[test]
fn dict_setitem_new_key() {
    assert_eq!(
        run_python_one("d = {}\nd['a'] = 1\nprint(d)\n"),
        "{'a': 1}"
    );
}

#[test]
fn dict_update_merge() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd.update({'b': 2})\nprint(d)\n"),
        "{'a': 1, 'b': 2}"
    );
}

#[test]
fn dict_pop_existing() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\nprint(d.pop('a'))\n"),
        "1"
    );
}

#[test]
fn dict_popitem_removes_arbitrary() {
    assert_eq!(
        run_python_one("d = {'only': 1}\nprint(d.popitem())\n"),
        "('only', 1)"
    );
}

#[test]
fn dict_keys_view() {
    assert_eq!(run_print("sorted({'b': 2, 'a': 1}.keys())"), "['a', 'b']");
}

#[test]
fn dict_values_view() {
    assert_eq!(run_print("sorted({'x': 3, 'y': 1}.values())"), "[1, 3]");
}

#[test]
fn dict_items_view() {
    assert_eq!(
        run_print("sorted({'a': 1}.items())"),
        "[('a', 1)]"
    );
}

#[test]
fn dict_in_keys() {
    assert_eq!(run_print("'a' in {'a': 1}"), "True");
}

#[test]
fn dict_not_in_missing_key() {
    assert_eq!(run_print("'z' in {'a': 1}"), "False");
}

#[test]
fn dict_len() {
    assert_eq!(run_print("len({'a': 1, 'b': 2})"), "2");
}

#[test]
fn dict_copy_shallow() {
    assert_eq!(
        run_python_one("d = {'a': 1}\ne = d.copy()\ne['b'] = 2\nprint(d, e)\n"),
        "{'a': 1} {'a': 1, 'b': 2}"
    );
}

#[test]
fn dict_clear() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd.clear()\nprint(d)\n"),
        "{}"
    );
}

#[test]
fn dict_from_pairs() {
    assert_eq!(run_print("dict([('a', 1), ('b', 2)])"), "{'a': 1, 'b': 2}");
}

#[test]
fn dict_from_keys() {
    assert_eq!(run_print("dict.fromkeys(['a', 'b'], 0)"), "{'a': 0, 'b': 0}");
}

#[test]
fn dict_setdefault_inserts() {
    assert_eq!(
        run_python_one("d = {}\nprint(d.setdefault('k', 5))\nprint(d)\n"),
        "5\n{'k': 5}"
    );
}

#[test]
fn dict_setdefault_existing() {
    assert_eq!(
        run_python_one("d = {'k': 1}\nprint(d.setdefault('k', 9))\n"),
        "1"
    );
}

#[test]
fn dict_merge_operator() {
    assert_eq!(run_print("{'a': 1} | {'b': 2}"), "{'a': 1, 'b': 2}");
}

#[test]
fn dict_merge_override() {
    assert_eq!(run_print("{'a': 1} | {'a': 2}"), "{'a': 2}");
}

#[test]
fn dict_nested_access() {
    assert_eq!(run_print("{'outer': {'inner': 3}}['outer']['inner']"), "3");
}

#[test]
fn dict_bool_empty_false() {
    assert_eq!(run_print("bool({})"), "False");
}

#[test]
fn dict_bool_nonempty_true() {
    assert_eq!(run_print("bool({'a': 1})"), "True");
}

#[test]
fn dict_equality() {
    assert_eq!(run_print("{'a': 1} == {'a': 1}"), "True");
}

#[test]
fn dict_inequality() {
    assert_eq!(run_print("{'a': 1} == {'a': 2}"), "False");
}

#[test]
fn dict_del_key() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\ndel d['a']\nprint(d)\n"),
        "{'b': 2}"
    );
}

#[test]
fn dict_pop_missing_with_default() {
    assert_eq!(run_print("{}.pop('x', None)"), "None");
}

#[test]
fn dict_comprehension_inline() {
    assert_eq!(run_print("{k: k for k in range(2)}"), "{0: 0, 1: 1}");
}

#[test]
fn dict_unpack_in_literal() {
    assert_eq!(run_print("{**{'a': 1}, **{'b': 2}}"), "{'a': 1, 'b': 2}");
}

#[test]
fn dict_iterate_keys_sum_values() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\nprint(sum(d[k] for k in d))\n"),
        "3"
    );
}

#[test]
fn dict_values_list_materialize() {
    assert_eq!(run_print("list({'x': 1, 'y': 2}.values())"), "[1, 2]");
}

#[test]
fn dict_keys_membership_loop() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\nprint('a' in d, 'c' in d)\n"),
        "True False"
    );
}

#[test]
fn dict_reversed_keys_insertion_order() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2, 'c': 3}\nprint(list(reversed(d)))\n"),
        "['c', 'b', 'a']"
    );
}

#[test]
fn dict_getitem_bracket() {
    assert_eq!(run_print("{'k': 7}['k']"), "7");
}

#[test]
fn dict_key_error_on_missing() {
    assert_eq!(
        run_python_one("try:\n {}['x']\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn dict_tuple_key() {
    assert_eq!(run_print("{(1, 2): 'pair'}[(1, 2)]"), "pair");
}

#[test]
fn dict_bool_key() {
    assert_eq!(run_print("{True: 'yes'}[True]"), "yes");
}

#[test]
fn dict_int_keys() {
    assert_eq!(run_print("{1: 'one', 2: 'two'}[2]"), "two");
}

#[test]
fn dict_none_value_allowed() {
    assert_eq!(run_print("{'k': None}['k']"), "None");
}

#[test]
fn dict_list_value_mutate() {
    assert_eq!(
        run_python_one("d = {'a': [1]}\nd['a'].append(2)\nprint(d)\n"),
        "{'a': [1, 2]}"
    );
}

#[test]
fn dict_sort_by_key() {
    assert_eq!(
        run_python_one("pairs = {'b': 2, 'a': 1}\nprint(sorted(pairs))\n"),
        "['a', 'b']"
    );
}

#[test]
fn dict_max_key_by_value() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 3, 'c': 2}\nprint(max(d, key=d.get))\n"),
        "b"
    );
}

#[test]
fn dict_items_unpack_in_for() {
    assert_eq!(
        run_python_one("total = 0\nfor k, v in {'a': 1, 'b': 2}.items():\n total += v\nprint(total)\n"),
        "3"
    );
}
