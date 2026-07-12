use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// List runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn list_literal_len() {
    assert_eq!(run_python_one("x = [1, 2, 3]\nprint(len(x))\n"), "3");
}

#[test]
fn list_empty() {
    assert_eq!(run_python_one("x = []\nprint(len(x))\n"), "0");
}

#[test]
fn list_index_access() {
    assert_eq!(run_python_one("x = [10, 20, 30]\nprint(x[1])\n"), "20");
}

#[test]
fn list_negative_index() {
    assert_eq!(run_python_one("x = [1, 2, 3]\nprint(x[-1])\n"), "3");
}

#[test]
fn list_append_runtime() {
    assert_eq!(
        run_python_one("x = [1, 2, 3]\nx.append(4)\nprint(len(x))\n"),
        "4"
    );
}

#[test]
fn list_nested() {
    compile_ok("x = [[1, 2], [3, 4]]\n");
}

#[test]
fn list_mixed_types() {
    compile_ok("x = [1, 'hello', True, None, 3.14]\n");
}

#[test]
fn list_trailing_comma() {
    parse_ok("x = [1, 2, 3,]\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Tuple runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn tuple_index() {
    assert_eq!(run_python_one("x = (10, 20)\nprint(x[0])\n"), "10");
}

#[test]
fn tuple_empty() {
    compile_ok("x = ()\n");
}

#[test]
fn tuple_single_element() {
    compile_ok("x = (1,)\n");
}

#[test]
fn tuple_trailing_comma() {
    parse_ok("x = (1, 2,)\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Dict runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn dict_access() {
    assert_eq!(run_python_one("x = {'a': 1, 'b': 2}\nprint(x['a'])\n"), "1");
}

#[test]
fn dict_empty() {
    assert_eq!(run_python_one("x = {}\nprint(len(x))\n"), "0");
}

#[test]
fn dict_set_key() {
    assert_eq!(
        run_python_one("d = {}\nd['key'] = 'value'\nprint(d['key'])\n"),
        "value"
    );
}

#[test]
fn dict_nested() {
    compile_ok("x = {'a': {'b': 1}}\n");
}

#[test]
fn dict_mixed_values() {
    compile_ok("x = {'a': [1, 2], 'b': [3, 4]}\n");
}

#[test]
fn dict_trailing_comma() {
    parse_ok("x = {'a': 1, 'b': 2,}\n");
}

#[test]
fn dict_unpacking() {
    compile_ok("a = {'x': 1}\nb = {'y': 2}\nc = {**a, **b}\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Set runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn set_literal() {
    compile_ok("x = {1, 2, 3}\n");
}

#[test]
fn set_with_constructor() {
    compile_ok("x = set()\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Slicing runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn slice_basic() {
    compile_ok("x = [1, 2, 3, 4, 5]\ny = x[1:3]\n");
}

#[test]
fn slice_step() {
    compile_ok("x = [1, 2, 3, 4, 5]\ny = x[::2]\n");
}

#[test]
fn slice_reverse() {
    compile_ok("x = [1, 2, 3, 4, 5]\ny = x[::-1]\n");
}

#[test]
fn slice_negative() {
    compile_ok("x = [1, 2, 3, 4, 5]\ny = x[-2:]\n");
}

#[test]
fn slice_negative_range() {
    compile_ok("x = [1, 2, 3, 4, 5]\ny = x[-3:-1]\n");
}

#[test]
fn string_slicing() {
    compile_ok("s = 'hello'[1:4]\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Assignment patterns
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn chained_assignment() {
    let out = run_python("a = b = c = 42\nprint(a)\nprint(b)\nprint(c)\n");
    assert_eq!(out[0], "42");
    assert_eq!(out[1], "42");
    assert_eq!(out[2], "42");
}

#[test]
fn subscript_assignment() {
    assert_eq!(
        run_python_one("x = [0, 0, 0]\nx[1] = 42\nprint(x[1])\n"),
        "42"
    );
}

#[test]
fn unpacking_starred() {
    compile_ok("first, *rest = [1, 2, 3, 4, 5]\n");
}

#[test]
fn unpacking_starred_middle() {
    compile_ok("first, *middle, last = [1, 2, 3, 4, 5]\n");
}

#[test]
fn unpacking_nested() {
    compile_ok("(a, b), c = (1, 2), 3\n");
}

#[test]
fn unpacking_list() {
    compile_ok("[a, [b, c]] = [1, [2, 3]]\n");
}
