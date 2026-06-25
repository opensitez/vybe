use crate::helpers::{run_print, run_python_one};

#[test]
fn json_dumps_list_of_ints() {
    assert_eq!(run_print("import json; json.dumps([1, 2, 3])"), "[1, 2, 3]");
}

#[test]
fn json_loads_list_roundtrip() {
    assert_eq!(
        run_print("import json; json.loads('[1, 2, 3]')"),
        "[1, 2, 3]"
    );
}

#[test]
fn json_dumps_dict_string_keys() {
    assert_eq!(
        run_print("import json; json.dumps({'a': 1, 'b': 2})"),
        "{\"a\": 1, \"b\": 2}"
    );
}

#[test]
fn json_loads_dict_roundtrip() {
    assert_eq!(
        run_python_one("import json\ns = json.dumps({'x': 9})\nprint(json.loads(s)['x'])\n"),
        "9"
    );
}

#[test]
fn json_dumps_bool_true() {
    assert_eq!(run_print("import json; json.dumps(True)"), "true");
}

#[test]
fn json_dumps_bool_false() {
    assert_eq!(run_print("import json; json.dumps(False)"), "false");
}

#[test]
fn json_dumps_null() {
    assert_eq!(run_print("import json; json.dumps(None)"), "null");
}

#[test]
fn json_loads_null_becomes_none() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('null'))\n"),
        "None"
    );
}

#[test]
fn json_dumps_string_quotes() {
    assert_eq!(run_print("import json; json.dumps('hi')"), "\"hi\"");
}

#[test]
fn json_loads_string() {
    assert_eq!(run_print("import json; json.loads('\"hi\"')"), "hi");
}

#[test]
fn json_dumps_nested_structure() {
    assert_eq!(
        run_print("import json; json.dumps({'items': [1, 2]})"),
        "{\"items\": [1, 2]}"
    );
}

#[test]
fn json_loads_nested_structure() {
    assert_eq!(
        run_python_one("import json\nd = json.loads('{\"items\": [1, 2]}')\nprint(d['items'][1])\n"),
        "2"
    );
}

#[test]
fn json_dumps_empty_list() {
    assert_eq!(run_print("import json; json.dumps([])"), "[]");
}

#[test]
fn json_dumps_empty_dict() {
    assert_eq!(run_print("import json; json.dumps({})"), "{}");
}

#[test]
fn json_loads_empty_list() {
    assert_eq!(run_print("import json; json.loads('[]')"), "[]");
}

#[test]
fn json_loads_empty_object() {
    assert_eq!(run_print("import json; json.loads('{}')"), "{}");
}

#[test]
fn json_dumps_float() {
    assert_eq!(run_print("import json; json.dumps(1.5)"), "1.5");
}

#[test]
fn json_loads_float() {
    assert_eq!(run_print("import json; json.loads('2.5')"), "2.5");
}

#[test]
fn json_roundtrip_tuple_becomes_list() {
    assert_eq!(
        run_python_one("import json\ns = json.dumps([1, 2])\nprint(type(json.loads(s)).__name__)\n"),
        "list"
    );
}

#[test]
fn json_dumps_with_indent_adds_newlines() {
    assert_eq!(
        run_python_one("import json\ns = json.dumps({'a': 1}, indent=2)\nprint('\\n' in s)\n"),
        "True"
    );
}

#[test]
fn json_loads_array_of_strings() {
    assert_eq!(
        run_print("import json; json.loads('[\"a\", \"b\"]')"),
        "['a', 'b']"
    );
}

#[test]
fn json_dumps_unicode_string() {
    assert_eq!(
        run_print("import json; json.dumps('é')"),
        "\"é\""
    );
}

#[test]
fn json_loads_unicode_escape() {
    assert_eq!(
        run_print("import json; json.loads('\"\\\\u00e9\"')"),
        "é"
    );
}

#[test]
fn json_dumps_sort_keys() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps({'b': 2, 'a': 1}, sort_keys=True))\n"),
        "{\"a\": 1, \"b\": 2}"
    );
}

#[test]
fn json_loads_integer_string() {
    assert_eq!(run_print("import json; json.loads('42')"), "42");
}

#[test]
fn json_dumps_negative_int() {
    assert_eq!(run_print("import json; json.dumps(-7)"), "-7");
}

#[test]
fn json_loads_negative_int() {
    assert_eq!(run_print("import json; json.loads('-7')"), "-7");
}

#[test]
fn json_dumps_list_of_bools() {
    assert_eq!(
        run_print("import json; json.dumps([True, False])"),
        "[true, false]"
    );
}

#[test]
fn json_loads_list_of_bools() {
    assert_eq!(
        run_print("import json; json.loads('[true, false]')"),
        "[True, False]"
    );
}

#[test]
fn json_double_roundtrip_stable() {
    assert_eq!(
        run_python_one("import json\ns = json.dumps({'k': [1]})\nprint(json.dumps(json.loads(s)))\n"),
        "{\"k\": [1]}"
    );
}

#[test]
fn json_loads_whitespace_padding() {
    assert_eq!(
        run_print("import json; json.loads('  [1]  ')"),
        "[1]"
    );
}

#[test]
fn json_dumps_special_chars_escaped() {
    assert_eq!(
        run_python_one("import json\ns = json.dumps('a\\nb')\nprint('\\\\n' in s)\n"),
        "True"
    );
}

#[test]
fn json_loads_object_with_null_field() {
    assert_eq!(
        run_python_one("import json\nd = json.loads('{\"x\": null}')\nprint(d['x'])\n"),
        "None"
    );
}

#[test]
fn json_dumps_deeply_nested() {
    assert_eq!(
        run_print("import json; json.dumps({'a': {'b': {'c': 1}}})"),
        "{\"a\": {\"b\": {\"c\": 1}}}"
    );
}

#[test]
fn json_loads_array_of_arrays() {
    assert_eq!(
        run_print("import json; json.loads('[[1], [2, 3]]')"),
        "[[1], [2, 3]]"
    );
}

#[test]
fn json_dumps_mixed_numeric_types() {
    assert_eq!(
        run_print("import json; json.dumps([1, 2.5, -3])"),
        "[1, 2.5, -3]"
    );
}

#[test]
fn json_loads_boolean_in_object() {
    assert_eq!(
        run_python_one("import json\nd = json.loads('{\"ok\": true}')\nprint(d['ok'])\n"),
        "True"
    );
}

#[test]
fn json_dumps_empty_string() {
    assert_eq!(run_print("import json; json.dumps('')"), "\"\"");
}

#[test]
fn json_loads_empty_string() {
    assert_eq!(run_print("import json; json.loads('\"\"')"), "");
}

#[test]
fn json_list_comprehension_roundtrip() {
    assert_eq!(
        run_python_one("import json\nvals = [json.loads(json.dumps(x)) for x in [1, 'a', None]]\nprint(vals)\n"),
        "[1, 'a', None]"
    );
}
