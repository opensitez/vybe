use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// Escape sequences
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn escape_newline() {
    let out = run_python_one("print('hello\\nworld')\n");
    assert_eq!(out, "hello\nworld");
}

#[test]
fn escape_tab() {
    assert_eq!(run_python_one("print('a\\tb')\n"), "a\tb");
}

#[test]
fn escape_backslash() {
    assert_eq!(run_python_one("print('a\\\\b')\n"), "a\\b");
}

#[test]
fn escape_single_quote() {
    assert_eq!(run_python_one("print('it\\'s')\n"), "it's");
}

#[test]
fn escape_double_quote() {
    assert_eq!(
        run_python_one("print(\"he said \\\"hi\\\"\")\n"),
        "he said \"hi\""
    );
}

#[test]
fn hex_escape_parse() {
    parse_ok("x = '\\x41'\n");
}

#[test]
fn unicode_escape_parse() {
    parse_ok("x = '\\u0041'\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Raw strings
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn raw_string_double() {
    compile_ok("x = r\"hello\\nworld\"\n");
}

#[test]
fn raw_string_single() {
    compile_ok("x = r'hello\\nworld'\n");
}

#[test]
fn raw_string_uppercase() {
    compile_ok("x = R\"no escapes here\"\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Byte strings
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn byte_string_lowercase() {
    parse_ok("x = b'hello'\n");
}

#[test]
fn byte_string_uppercase() {
    parse_ok("x = B'hello'\n");
}

#[test]
fn byte_string_with_escapes() {
    parse_ok("x = b\"\\x48\\x65\\x6c\\x6c\\x6f\"\n");
}

#[test]
fn raw_byte_string_rb() {
    parse_ok("x = rb'raw bytes'\n");
}

#[test]
fn raw_byte_string_br() {
    parse_ok("x = br'raw bytes'\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// Triple-quoted strings
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn triple_double_with_embedded_quotes() {
    parse_ok("x = \"\"\"He said \"hello\" to her\"\"\"\n");
}

#[test]
fn triple_single_with_embedded_quotes() {
    parse_ok("x = '''It's a 'test' string'''\n");
}

#[test]
fn triple_multiline() {
    compile_ok("x = \"\"\"line1\nline2\nline3\"\"\"\n");
}

#[test]
fn triple_single_multiline() {
    compile_ok("x = '''first\nsecond\nthird'''\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// F-string runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn fstring_basic_runtime() {
    assert_eq!(
        run_python_one("name = 'world'\nprint(f'hello {name}')\n"),
        "hello world"
    );
}

#[test]
fn fstring_expression() {
    assert_eq!(run_python_one("x = 3\nprint(f'{x + 1}')\n"), "4");
}

#[test]
fn fstring_multiple_exprs() {
    assert_eq!(
        run_python_one("a = 1\nb = 2\nprint(f'{a} + {b} = {a + b}')\n"),
        "1 + 2 = 3"
    );
}

#[test]
fn fstring_escaped_braces() {
    assert_eq!(run_python_one("print(f'{{hello}}')\n"), "{hello}");
}

#[test]
fn fstring_nested_quotes() {
    assert_eq!(
        run_python_one("d = {'key': 'val'}\nprint(f\"{d['key']}\")\n"),
        "val"
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// String methods runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn str_upper() {
    assert_eq!(run_python_one("print('hello'.upper())\n"), "HELLO");
}

#[test]
fn str_lower() {
    assert_eq!(run_python_one("print('HELLO'.lower())\n"), "hello");
}

#[test]
fn str_strip() {
    assert_eq!(run_python_one("print('  hi  '.strip())\n"), "hi");
}

#[test]
fn str_replace_runtime() {
    assert_eq!(
        run_python_one("print('hello'.replace('l', 'r'))\n"),
        "herro"
    );
}

#[test]
fn str_startswith() {
    assert_eq!(run_python_one("print('hello'.startswith('he'))\n"), "true");
}

#[test]
fn str_endswith() {
    assert_eq!(run_python_one("print('hello'.endswith('lo'))\n"), "true");
}

#[test]
fn str_find() {
    assert_eq!(run_python_one("print('hello'.find('lo'))\n"), "3");
}

#[test]
fn str_split_runtime() {
    let out = run_python("parts = 'a,b,c'.split(',')\nfor p in parts:\n    print(p)\n");
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn str_join_runtime() {
    assert_eq!(
        run_python_one("print(', '.join(['a', 'b', 'c']))\n"),
        "a, b, c"
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// String multiplication
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn str_multiply() {
    assert_eq!(run_python_one("print('ab' * 3)\n"), "ababab");
}

#[test]
fn str_multiply_reverse() {
    assert_eq!(run_python_one("print(3 * 'ab')\n"), "ababab");
}

// ══════════════════════════════════════════════════════════════════════════════
// Numeric literals
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn hex_literal() {
    parse_ok("x = 0xFF\n");
}

#[test]
fn octal_literal() {
    parse_ok("y = 0o77\n");
}

#[test]
fn binary_literal() {
    parse_ok("z = 0b1010\n");
}

#[test]
fn underscore_in_number() {
    parse_ok("x = 1_000_000\n");
}

#[test]
fn underscore_in_float() {
    parse_ok("y = 3.14_15_93\n");
}

#[test]
fn hex_underscore() {
    parse_ok("w = 0xFF_FF\n");
}

#[test]
fn complex_literal() {
    parse_ok("x = 3+4j\n");
}

#[test]
fn complex_literal_upper() {
    parse_ok("y = 1J\n");
}

#[test]
fn complex_literal_float() {
    parse_ok("z = .5j\n");
}

#[test]
fn ellipsis_literal() {
    parse_ok("x = ...\n");
}

#[test]
fn ellipsis_in_function() {
    compile_ok("def stub():\n    ...\n");
}
