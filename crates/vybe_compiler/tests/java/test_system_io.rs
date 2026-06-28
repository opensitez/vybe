use crate::helpers::run_main;

#[test]
fn println_prints_string_literal() {
    let out = run_main(r#"System.out.println("hello");"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn println_prints_integer_value() {
    let out = run_main("System.out.println(42);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn println_prints_boolean_true() {
    let out = run_main("System.out.println(true);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn println_prints_boolean_false() {
    let out = run_main("System.out.println(false);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn println_prints_null_reference() {
    let out = run_main("String s = null; System.out.println(s);");
    assert_eq!(out, vec!["null"]);
}

#[test]
fn println_multiple_calls_emit_separate_lines() {
    let out = run_main("System.out.println(1); System.out.println(2); System.out.println(3);");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn println_concatenates_string_and_integer() {
    let out = run_main(r#"System.out.println("count=" + 7);"#);
    assert_eq!(out, vec!["count=7"]);
}

#[test]
fn println_with_no_arguments_prints_blank_line() {
    let out = run_main("System.out.println();");
    assert_eq!(out, vec![""]);
}

#[test]
fn print_writes_without_extra_line_from_single_call() {
    let out = run_main(r#"System.out.print("partial");"#);
    assert_eq!(out, vec!["partial"]);
}

#[test]
fn print_followed_by_println_combines_on_one_line() {
    let out = run_main(r#"System.out.print("ab"); System.out.println("c");"#);
    assert_eq!(out, vec!["ab", "abc"]);
}

#[test]
fn print_integer_without_newline() {
    let out = run_main("System.out.print(9);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn print_then_println_with_strings() {
    let out = run_main(r#"System.out.print("foo"); System.out.println("bar");"#);
    assert_eq!(out, vec!["foo", "foobar"]);
}

#[test]
fn printf_formats_integer_placeholder() {
    let out = run_main(r#"System.out.printf("%d", 15);"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn printf_formats_string_placeholder() {
    let out = run_main(r#"System.out.printf("%s", "vybe");"#);
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn printf_formats_labeled_integer_with_text() {
    let out = run_main(r#"System.out.printf("n=%d", 8);"#);
    assert_eq!(out, vec!["n=8"]);
}

#[test]
fn printf_formats_multiple_placeholders_in_order() {
    let out = run_main(r#"System.out.printf("%s:%d", "id", 3);"#);
    assert_eq!(out, vec!["id:3"]);
}

#[test]
fn printf_percent_literal_doubles_percent_sign() {
    let out = run_main(r#"System.out.printf("100%%");"#);
    assert_eq!(out, vec!["100%"]);
}

#[test]
fn printf_newline_escape_via_percent_n() {
    let out = run_main(r#"System.out.printf("a%nb");"#);
    assert_eq!(out, vec!["a\nb"]);
}

#[test]
fn printf_hex_lowercase_placeholder() {
    let out = run_main(r#"System.out.printf("%x", 255);"#);
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn printf_hex_uppercase_placeholder() {
    let out = run_main(r#"System.out.printf("%X", 255);"#);
    assert_eq!(out, vec!["FF"]);
}

#[test]
fn get_property_missing_key_is_null() {
    let out = run_main(r#"System.out.println(System.getProperty("vybe.missing.property.key"));"#);
    assert_eq!(out, vec!["null"]);
}

#[test]
fn get_property_missing_key_equals_null() {
    let out = run_main(r#"System.out.println(System.getProperty("vybe.absent.key") == null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn get_property_result_can_be_coalesced_with_ternary() {
    let out = run_main(
        r#"String v = System.getProperty("vybe.absent.key"); String out = v != null ? v : "default"; System.out.println(out);"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn get_property_on_known_style_key_returns_string_or_null() {
    let out = run_main(
        r#"String v = System.getProperty("line.separator"); System.out.println(v != null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn println_printf_and_print_can_mix_in_sequence() {
    let out =
        run_main(r#"System.out.print("["); System.out.printf("%d", 5); System.out.println("]");"#);
    assert_eq!(out, vec!["[", "5", "5]"]);
}
