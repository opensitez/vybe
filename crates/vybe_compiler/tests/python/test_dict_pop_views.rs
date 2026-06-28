use crate::helpers::{run_print, run_python_one};

#[test]
fn dict_popitem_removes_arbitrary_pair() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\nk, v = d.popitem()\nprint(len(d))\n"),
        "1"
    );
}

#[test]
fn dict_pop_existing_key() {
    assert_eq!(
        run_python_one("d = {'x': 10, 'y': 20}\nprint(d.pop('x'))\n"),
        "10"
    );
}

#[test]
fn dict_pop_missing_with_default() {
    assert_eq!(run_python_one("d = {}\nprint(d.pop('z', 99))\n"), "99");
}

#[test]
fn dict_keys_view_length() {
    assert_eq!(run_print("len({'a': 1, 'b': 2}.keys())"), "2");
}

#[test]
fn dict_values_view_contains() {
    assert_eq!(run_print("2 in {'a': 1, 'b': 2}.values()"), "True");
}

#[test]
fn dict_items_view_to_list() {
    assert_eq!(
        run_print("sorted({'b': 2, 'a': 1}.items())"),
        "[('a', 1), ('b', 2)]"
    );
}

#[test]
fn dict_keys_iteration_order_insertion() {
    assert_eq!(
        run_python_one("d = {}\nd['z'] = 1\nd['a'] = 2\nprint(list(d.keys())[0])\n"),
        "z"
    );
}

#[test]
fn dict_update_overwrites() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nd.update({'a': 9, 'b': 2})\nprint(d['a'], d['b'])\n"),
        "9 2"
    );
}

#[test]
fn dict_setdefault_inserts_missing() {
    assert_eq!(
        run_python_one("d = {}\nprint(d.setdefault('k', 5))\nprint(d['k'])\n"),
        "5"
    );
}

#[test]
fn dict_setdefault_keeps_existing() {
    assert_eq!(
        run_python_one("d = {'k': 1}\nprint(d.setdefault('k', 5))\n"),
        "1"
    );
}

#[test]
fn dict_clear_empties() {
    assert_eq!(run_python_one("d = {'a': 1}\nd.clear()\nprint(d)\n"), "{}");
}

#[test]
fn dict_copy_is_shallow() {
    assert_eq!(
        run_python_one("a = {'x': [1]}\nb = a.copy()\nb['x'].append(2)\nprint(a['x'])\n"),
        "[1, 2]"
    );
}

#[test]
fn dict_fromkeys_default_none() {
    assert_eq!(
        run_print("dict.fromkeys(['a', 'b'])"),
        "{'a': None, 'b': None}"
    );
}

#[test]
fn dict_fromkeys_custom_value() {
    assert_eq!(
        run_print("dict.fromkeys(['a', 'b'], 0)"),
        "{'a': 0, 'b': 0}"
    );
}

#[test]
fn dict_get_missing_default() {
    assert_eq!(run_print("{}.get('missing', 'fallback')"), "fallback");
}

#[test]
fn dict_get_existing_no_default() {
    assert_eq!(run_print("{'a': 1}.get('a')"), "1");
}

#[test]
fn dict_popitem_on_single_item() {
    assert_eq!(
        run_python_one("d = {'only': 1}\nd.popitem()\nprint(d)\n"),
        "{}"
    );
}

#[test]
fn dict_items_unpack_in_for_loop() {
    assert_eq!(
        run_python_one(
            "total = 0\nfor k, v in {'a': 1, 'b': 2}.items():\n total += v\nprint(total)\n"
        ),
        "3"
    );
}

#[test]
fn dict_keys_membership() {
    assert_eq!(run_print("'a' in {'a': 1, 'b': 2}"), "True");
}

#[test]
fn dict_values_membership_false_for_key() {
    assert_eq!(run_print("'a' in {'a': 1}.values()"), "False");
}

#[test]
fn dict_del_key() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\ndel d['a']\nprint('a' in d)\n"),
        "False"
    );
}

#[test]
fn dict_merge_union_operator() {
    assert_eq!(
        run_python_one("a = {'x': 1}\nb = {'y': 2}\nc = a | b\nprint(sorted(c.items()))\n"),
        "[('x', 1), ('y', 2)]"
    );
}

#[test]
fn dict_inplace_union_update() {
    assert_eq!(
        run_python_one("a = {'x': 1}\na |= {'x': 9, 'z': 3}\nprint(a['x'], a['z'])\n"),
        "9 3"
    );
}

#[test]
fn dict_comprehension_filter_items() {
    assert_eq!(
        run_print("{k: v for k, v in {'a': 1, 'b': 2, 'c': 3}.items() if v > 1}"),
        "{'b': 2, 'c': 3}"
    );
}

#[test]
fn dict_nested_access() {
    assert_eq!(run_print("{'outer': {'inner': 7}}['outer']['inner']"), "7");
}

#[test]
fn dict_values_list_sum() {
    assert_eq!(run_print("sum({'a': 1, 'b': 2, 'c': 3}.values())"), "6");
}

#[test]
fn dict_keys_list_sorted() {
    assert_eq!(
        run_print("sorted({'z': 1, 'a': 2, 'm': 3}.keys())"),
        "['a', 'm', 'z']"
    );
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
fn dict_popitem_return_is_tuple() {
    assert_eq!(
        run_python_one("d = {'k': 9}\nitem = d.popitem()\nprint(len(item))\n"),
        "2"
    );
}

#[test]
fn dict_equality_same_content() {
    assert_eq!(run_print("{'a': 1} == {'a': 1}"), "True");
}

#[test]
fn dict_inequality_different_keys() {
    assert_eq!(run_print("{'a': 1} != {'b': 1}"), "True");
}

#[test]
fn dict_repr_empty() {
    assert_eq!(run_print("repr({})"), "{}");
}

#[test]
fn dict_len_counts_pairs() {
    assert_eq!(run_print("len({'a': 1, 'b': 2, 'c': 3})"), "3");
}

#[test]
fn dict_update_with_kwargs_style_dict() {
    assert_eq!(
        run_python_one("d = {}\nd.update({'a': 1})\nd.update({'b': 2})\nprint(len(d))\n"),
        "2"
    );
}

#[test]
fn dict_items_not_same_as_keys() {
    assert_eq!(
        run_python_one("d = {'a': 1}\nprint(list(d.items()) == list(d.keys()))\n"),
        "False"
    );
}

#[test]
fn dict_pop_removes_from_len() {
    assert_eq!(
        run_python_one("d = {'a': 1, 'b': 2}\nd.pop('a')\nprint(len(d))\n"),
        "1"
    );
}

#[test]
fn dict_setdefault_return_on_insert() {
    assert_eq!(
        run_python_one("d = {}\nr = d.setdefault('x', [1])\nr.append(2)\nprint(d['x'])\n"),
        "[1, 2]"
    );
}

#[test]
fn dict_view_iterable_in_list() {
    assert_eq!(run_print("list({'a': 1, 'b': 2}.values())"), "[1, 2]");
}
