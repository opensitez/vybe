use crate::helpers::{run_python_one, run_print};

#[test]
fn json_dumps_dict_sorted_keys_style() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps({'b': 2, 'a': 1}, sort_keys=True))\n"),
        "{\"a\": 1, \"b\": 2}"
    );
}

#[test]
fn json_loads_object() {
    assert_eq!(
        run_python_one("import json\nd = json.loads('{\"x\": 3}')\nprint(d['x'])\n"),
        "3"
    );
}

#[test]
fn json_loads_array() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('[1, 2, 3]')[1])\n"),
        "2"
    );
}

#[test]
fn json_dumps_list() {
    assert_eq!(run_print("import json; json.dumps([1, 2])"), "[1, 2]");
}

#[test]
fn json_roundtrip_string() {
    assert_eq!(
        run_python_one("import json\ns = 'hello'\nprint(json.loads(json.dumps(s)))\n"),
        "hello"
    );
}

#[test]
fn json_dumps_bool_true() {
    assert_eq!(run_print("import json; json.dumps(True)"), "true");
}

#[test]
fn json_dumps_null_none() {
    assert_eq!(run_print("import json; json.dumps(None)"), "null");
}

#[test]
fn json_loads_null_to_none() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('null'))\n"),
        "None"
    );
}

#[test]
fn json_dumps_nested() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps({'a': {'b': 1}}))\n"),
        "{\"a\": {\"b\": 1}}"
    );
}

#[test]
fn json_loads_number_int() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('42'))\n"),
        "42"
    );
}

#[test]
fn json_loads_number_float() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('3.5'))\n"),
        "3.5"
    );
}

#[test]
fn re_search_finds_digits() {
    assert_eq!(
        run_python_one("import re\nm = re.search(r'\\d+', 'ab12cd')\nprint(m.group() if m else 'none')\n"),
        "12"
    );
}

#[test]
fn re_search_no_match() {
    assert_eq!(
        run_python_one("import re\nm = re.search(r'\\d+', 'abc')\nprint('none' if m is None else m.group())\n"),
        "none"
    );
}

#[test]
fn re_findall_digits() {
    assert_eq!(
        run_print("import re; re.findall(r'\\d+', 'a1b22c')"),
        "['1', '22']"
    );
}

#[test]
fn re_sub_replace_digits() {
    assert_eq!(
        run_python_one("import re\nprint(re.sub(r'\\d', 'X', 'a1b2'))\n"),
        "aXbX"
    );
}

#[test]
fn re_split_on_whitespace() {
    assert_eq!(
        run_print("import re; re.split(r'\\s+', 'a  b   c')"),
        "['a', 'b', 'c']"
    );
}

#[test]
fn re_match_start_of_string() {
    assert_eq!(
        run_python_one("import re\nm = re.match(r'abc', 'abcdef')\nprint(bool(m))\n"),
        "True"
    );
}

#[test]
fn re_match_fails_not_at_start() {
    assert_eq!(
        run_python_one("import re\nm = re.match(r'bc', 'abcdef')\nprint(m is None)\n"),
        "True"
    );
}

#[test]
fn re_findall_words() {
    assert_eq!(
        run_print("import re; re.findall(r'[a-z]+', 'Hello, world!')"),
        "['ello', 'world']"
    );
}

#[test]
fn re_sub_count_limit() {
    assert_eq!(
        run_python_one("import re\nprint(re.sub(r'a', 'b', 'aaa', count=2))\n"),
        "bba"
    );
}

#[test]
fn re_split_maxsplit() {
    assert_eq!(
        run_print("import re; re.split(r',', 'a,b,c,d', maxsplit=2)"),
        "['a', 'b', 'c,d']"
    );
}

#[test]
fn re_search_group_index() {
    assert_eq!(
        run_python_one("import re\nm = re.search(r'(\\d+)(\\d+)', 'ab1234')\nprint(m.group(2) if m else '')\n"),
        "34"
    );
}

#[test]
fn re_findall_capturing_groups_returns_tuples() {
    assert_eq!(
        run_print("import re; re.findall(r'(\\w)(\\w)', 'ab')"),
        "[('a', 'b')]"
    );
}

#[test]
fn json_dumps_unicode_string() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps('café'))\n"),
        "\"café\""
    );
}

#[test]
fn json_loads_empty_object() {
    assert_eq!(run_print("import json; json.loads('{}')"), "{}");
}

#[test]
fn json_loads_empty_array() {
    assert_eq!(run_print("import json; json.loads('[]')"), "[]");
}

#[test]
fn re_sub_backreference_style() {
    assert_eq!(
        run_python_one("import re\nprint(re.sub(r'(a)(b)', r'\\2\\1', 'ab'))\n"),
        "ba"
    );
}

#[test]
fn re_search_case_sensitive_default() {
    assert_eq!(
        run_python_one("import re\nprint(re.search(r'abc', 'xxABCxx') is None)\n"),
        "True"
    );
}

#[test]
fn re_findall_empty_pattern_matches() {
    assert_eq!(
        run_python_one("import re\nprint(len(re.findall(r'a', '')))\n"),
        "0"
    );
}

#[test]
fn json_dumps_escape_quotes() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps('say \"hi\"'))\n"),
        "\"say \\\"hi\\\"\""
    );
}

#[test]
fn re_split_capturing_groups_kept() {
    assert_eq!(
        run_print("import re; re.split(r'(:)', 'a:b')"),
        "['a', ':', 'b']"
    );
}

#[test]
fn json_loads_bool_true() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('true'))\n"),
        "True"
    );
}

#[test]
fn json_loads_bool_false() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('false'))\n"),
        "False"
    );
}

#[test]
fn re_findall_multichar_classes() {
    assert_eq!(
        run_print("import re; re.findall(r'[0-9]+', 'v1.2.3')"),
        "['1', '2', '3']"
    );
}

#[test]
fn re_sub_function_replacement() {
    assert_eq!(
        run_python_one("import re\nprint(re.sub(r'\\d+', lambda m: str(int(m.group()) * 2), 'a3b10'))\n"),
        "a6b20"
    );
}

#[test]
fn json_nested_roundtrip() {
    assert_eq!(
        run_python_one("import json\nobj = {'items': [1, {'k': 'v'}]}\nprint(json.loads(json.dumps(obj))['items'][1]['k'])\n"),
        "v"
    );
}

#[test]
fn re_match_digit_at_start() {
    assert_eq!(
        run_python_one("import re\nm = re.match(r'\\d+', '99abc')\nprint(m.group())\n"),
        "99"
    );
}

#[test]
fn re_search_dot_matches_any() {
    assert_eq!(
        run_python_one("import re\nm = re.search(r'a.c', 'abc')\nprint(m.group())\n"),
        "abc"
    );
}

#[test]
fn json_dumps_indent_pretty() {
    assert_eq!(
        run_python_one("import json\nout = json.dumps({'a': 1}, indent=2)\nprint('a' in out)\n"),
        "True"
    );
}

#[test]
fn re_findall_word_boundaries_simple() {
    assert_eq!(
        run_print("import re; re.findall(r'cat', 'concatenate cat')"),
        "['cat', 'cat']"
    );
}

#[test]
fn json_loads_string_with_backslash() {
    assert_eq!(
        run_python_one("import json\nprint(json.loads('\"a\\\\nb\"'))\n"),
        "a\nb"
    );
}

#[test]
fn re_split_on_comma_optional_space() {
    assert_eq!(
        run_print("import re; re.split(r',\\s*', 'a, b ,c')"),
        "['a', 'b', 'c']"
    );
}

#[test]
fn json_dumps_list_of_dicts() {
    assert_eq!(
        run_python_one("import json\nprint(json.dumps([{'id': 1}, {'id': 2}]))\n"),
        "[{\"id\": 1}, {\"id\": 2}]"
    );
}

#[test]
fn re_sub_removes_whitespace() {
    assert_eq!(
        run_python_one("import re\nprint(re.sub(r'\\s+', '', 'a b c'))\n"),
        "abc"
    );
}
