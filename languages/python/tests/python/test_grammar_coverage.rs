use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// 1. F-string conversion specs (!r, !s, !a)
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn fstring_conversion_r() {
    parse_ok("x = f\"{42!r}\"");
}

#[test]
fn fstring_conversion_s() {
    parse_ok("x = f\"{42!s}\"");
}

#[test]
fn fstring_conversion_a() {
    parse_ok("x = f\"{42!a}\"");
}

#[test]
fn fstring_conversion_with_format_spec() {
    parse_ok("x = f\"{val!r:.20}\"");
}

#[test]
fn fstring_conversion_runtime() {
    let out = run_python("x = 'hello'\nprint(f\"{x!s}\")\n");
    assert_eq!(out[0], "hello");
}

// ══════════════════════════════════════════════════════════════════════════════
// 2. Implicit string concatenation
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn string_concat_adjacent() {
    let out = run_python("x = \"hello\" \" world\"\nprint(x)\n");
    assert_eq!(out[0], "hello world");
}

#[test]
fn string_concat_multiple() {
    let out = run_python("x = \"a\" \"b\" \"c\"\nprint(x)\n");
    assert_eq!(out[0], "abc");
}

#[test]
fn string_concat_mixed_quotes() {
    let out = run_python("x = 'hello' \" world\"\nprint(x)\n");
    assert_eq!(out[0], "hello world");
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. u/U string prefix
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn u_string_lowercase() {
    let out = run_python("x = u\"hello\"\nprint(x)\n");
    assert_eq!(out[0], "hello");
}

#[test]
fn u_string_uppercase() {
    let out = run_python("x = U\"hello\"\nprint(x)\n");
    assert_eq!(out[0], "hello");
}

#[test]
fn u_string_single_quotes() {
    let out = run_python("x = u'world'\nprint(x)\n");
    assert_eq!(out[0], "world");
}

// ══════════════════════════════════════════════════════════════════════════════
// 4. Raw f-string prefixes (rf, fr, etc.)
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn raw_fstring_rf() {
    parse_ok("x = rf\"{42}\"");
}

#[test]
fn raw_fstring_fr() {
    parse_ok("x = fr\"{42}\"");
}

#[test]
fn raw_fstring_rf_upper() {
    parse_ok("x = RF\"{42}\"");
}

#[test]
fn raw_fstring_fr_mixed() {
    parse_ok("x = Fr\"{42}\"");
}

#[test]
fn raw_fstring_runtime() {
    let out = run_python("n = 5\nprint(rf\"{n}\")\n");
    assert_eq!(out[0], "5");
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. Triple-quoted f-strings (ordering bug fix)
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn fstring_triple_double() {
    let out = run_python("x = 10\nprint(f\"\"\"value is {x}\"\"\")\n");
    assert_eq!(out[0], "value is 10");
}

#[test]
fn fstring_triple_single() {
    let out = run_python("x = 10\nprint(f'''value is {x}''')\n");
    assert_eq!(out[0], "value is 10");
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. Pattern `as` bindings
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn match_as_pattern() {
    compile_ok(
        "match [1, 2]:\n    case [x, y] as point:\n        print(point)\n    case _:\n        pass\n",
    );
}

#[test]
fn match_as_wildcard() {
    compile_ok("match 42:\n    case x as val:\n        print(val)\n");
}

#[test]
fn match_as_singleton() {
    compile_ok("match True:\n    case True as v:\n        pass\n    case _:\n        pass\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// 7. Keyword patterns in class matching
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn match_class_keyword_pattern() {
    compile_ok(
        "class Point:\n    x = 0\n    y = 0\nmatch p:\n    case Point(x=1, y=2):\n        print('origin-ish')\n    case _:\n        pass\n",
    );
}

#[test]
fn match_class_mixed_positional_keyword() {
    compile_ok("match p:\n    case Rect(a, width=10):\n        pass\n    case _:\n        pass\n");
}

// ══════════════════════════════════════════════════════════════════════════════
// 8. match/case as soft keywords (usable as identifiers)
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn match_as_variable_name() {
    let out = run_python("match = 42\nprint(match)\n");
    assert_eq!(out[0], "42");
}

#[test]
fn case_as_variable_name() {
    let out = run_python("case = 'hello'\nprint(case)\n");
    assert_eq!(out[0], "hello");
}

#[test]
fn match_case_as_function_names() {
    let out = run_python(
        "def match(x):\n    return x + 1\ndef case(x):\n    return x * 2\nprint(match(5))\nprint(case(3))\n",
    );
    assert_eq!(out[0], "6");
    assert_eq!(out[1], "6");
}
