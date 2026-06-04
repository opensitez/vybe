use super::helpers::{compile_ok, run_prints};

// ── More string methods ──────────────────────────────────────

#[test]
fn ends_with() {
    compile_ok("var b = 'hello.dart'.endsWith('.dart');");
}
#[test]
fn ends_with_result() {
    let out = run_prints("void main() { print('hello.dart'.endsWith('.dart')); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn index_of() {
    compile_ok("var i = 'hello world'.indexOf('world');");
}
#[test]
fn index_of_result() {
    let out = run_prints("void main() { print('hello world'.indexOf('world')); }");
    assert_eq!(out, ["6"]);
}

#[test]
fn last_index_of() {
    compile_ok("var i = 'abcabc'.lastIndexOf('b');");
}
#[test]
fn last_index_of_result() {
    let out = run_prints("void main() { print('abcabc'.lastIndexOf('b')); }");
    assert_eq!(out, ["4"]);
}

#[test]
fn pad_right() {
    compile_ok("var s = '42'.padRight(5, '0');");
}
#[test]
fn pad_right_result() {
    let out = run_prints("void main() { print('hi'.padRight(5, '.')); }");
    assert_eq!(out, ["hi..."]);
}

#[test]
fn pad_left_result() {
    let out = run_prints("void main() { print('7'.padLeft(3, '0')); }");
    assert_eq!(out, ["007"]);
}

#[test]
fn trim_left() {
    compile_ok("var s = '  hello'.trimLeft();");
}
#[test]
fn trim_right() {
    compile_ok("var s = 'hello  '.trimRight();");
}

#[test]
fn compare_to() {
    compile_ok("var r = 'apple'.compareTo('banana');");
}
#[test]
fn compare_to_result() {
    let out = run_prints("void main() { print('a'.compareTo('a')); }");
    assert_eq!(out, ["0"]);
}

#[test]
fn is_empty() {
    compile_ok("var b = ''.isEmpty;");
}
#[test]
fn is_not_empty() {
    compile_ok("var b = 'x'.isNotEmpty;");
}

#[test]
fn is_empty_result() {
    let out = run_prints("void main() { print(''.isEmpty); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn is_not_empty_result() {
    let out = run_prints("void main() { print('x'.isNotEmpty); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn code_unit_at() {
    compile_ok("var c = 'A'.codeUnitAt(0);");
}
#[test]
fn replace_first() {
    compile_ok("var s = 'aaa'.replaceFirst('a', 'b');");
}

#[test]
fn replace_first_result() {
    let out = run_prints("void main() { print('aaa'.replaceFirst('a', 'b')); }");
    assert_eq!(out, ["baa"]);
}

#[test]
fn to_lower_result() {
    let out = run_prints("void main() { print('HELLO'.toLowerCase()); }");
    assert_eq!(out, ["hello"]);
}

#[test]
fn to_upper_result() {
    let out = run_prints("void main() { print('hello'.toUpperCase()); }");
    assert_eq!(out, ["HELLO"]);
}

#[test]
fn trim_result() {
    let out = run_prints("void main() { print('  hi  '.trim()); }");
    assert_eq!(out, ["hi"]);
}

#[test]
fn contains_result() {
    let out = run_prints("void main() { print('hello'.contains('ell')); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn split_result() {
    let out = run_prints("void main() { var parts = 'a,b,c'.split(','); print(parts.length); }");
    assert_eq!(out, ["3"]);
}

#[test]
fn starts_with_result() {
    let out = run_prints("void main() { print('hello'.startsWith('he')); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn substring_result() {
    let out = run_prints("void main() { print('hello'.substring(1, 3)); }");
    assert_eq!(out, ["el"]);
}

// ── Multi-line strings ───────────────────────────────────────

#[test]
fn multiline_single_quote() {
    compile_ok("var s = '''line 1\nline 2\nline 3''';");
}
#[test]
fn multiline_double_quote() {
    compile_ok("var s = \"\"\"line 1\nline 2\"\"\";");
}
#[test]
fn multiline_has_content() {
    compile_ok(
        r#"var text = '''
Hello
World
''';"#,
    );
}

// ── Raw strings ──────────────────────────────────────────────

#[test]
fn raw_string() {
    compile_ok("var s = r'C:\\Users\\test';");
}
#[test]
fn raw_string_regex() {
    compile_ok("var pat = r'\\d+\\.\\d+';");
}
#[test]
fn raw_string_no_interp() {
    compile_ok("var name = 'World'; var s = r'Hello $name';");
}

#[test]
fn raw_string_result() {
    let out = run_prints(r#"void main() { var s = r'no\n escape'; print(s.length > 5); }"#);
    assert_eq!(out, ["true"]);
}

// ── String interpolation (more cases) ───────────────────────

#[test]
fn interp_method_call() {
    compile_ok("var s = 'hello'; var r = 'upper: ${s.toUpperCase()}';");
}
#[test]
fn interp_arithmetic() {
    compile_ok("var n = 5; var s = 'result: ${n * n}';");
}
#[test]
fn interp_nested_class() {
    compile_ok(
        "class Dog { String name; Dog(this.name); } void main() { var d = Dog('Rex'); var s = 'Dog: ${d.name}'; }",
    );
}

#[test]
fn interp_result() {
    let out = run_prints("void main() { var x = 3; print('x is $x'); }");
    assert_eq!(out, ["x is 3"]);
}

#[test]
fn interp_expr_result() {
    let out = run_prints("void main() { var x = 3; print('squared: ${x * x}'); }");
    assert_eq!(out, ["squared: 9"]);
}

// ── toString ────────────────────────────────────────────────

#[test]
fn to_string_int() {
    compile_ok("var s = 42.toString();");
}
#[test]
fn to_string_double() {
    compile_ok("var s = 3.14.toString();");
}
#[test]
fn to_string_bool() {
    compile_ok("var s = true.toString();");
}

#[test]
fn to_string_result() {
    let out = run_prints("void main() { print(42.toString()); }");
    assert_eq!(out, ["42"]);
}

// ── String length ────────────────────────────────────────────

#[test]
fn string_length() {
    compile_ok("var n = 'hello'.length;");
}
#[test]
fn string_length_result() {
    let out = run_prints("void main() { print('hello'.length); }");
    assert_eq!(out, ["5"]);
}

// ── String joining ────────────────────────────────────────────

#[test]
fn join_strings() {
    let out = run_prints("void main() { var parts = ['a', 'b', 'c']; print(parts.join('-')); }");
    assert_eq!(out, ["a-b-c"]);
}

// ── Number → string conversions ──────────────────────────────

#[test]
fn int_parse() {
    compile_ok("var n = int.parse('42');");
}
#[test]
fn double_parse() {
    compile_ok("var n = double.parse('3.14');");
}

#[test]
fn int_parse_result() {
    let out = run_prints("void main() { print(int.parse('42') + 1); }");
    assert_eq!(out, ["43"]);
}
