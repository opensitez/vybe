use crate::helpers::run_main;

#[test]
fn string_concatenation_with_plus() {
    let out = run_main(r#"String s = "hello" + " " + "world"; System.out.println(s);"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn string_length_counts_code_units() {
    let out = run_main(r#"String s = "abcde"; System.out.println(s.length());"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_to_upper_and_lower_change_case() {
    let out = run_main(
        r#"String s = "Hello"; System.out.println(s.toUpperCase()); System.out.println(s.toLowerCase());"#,
    );
    assert_eq!(out, vec!["HELLO", "hello"]);
}

#[test]
fn string_substring_from_index() {
    let out = run_main(r#"String s = "foobar"; System.out.println(s.substring(3));"#);
    assert_eq!(out, vec!["bar"]);
}

#[test]
fn string_contains_reports_membership() {
    let out = run_main(
        r#"String s = "abcdef"; System.out.println(s.contains("cde")); System.out.println(s.contains("xyz"));"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_trim_strips_leading_and_trailing_whitespace() {
    let out = run_main(r#"String s = "  hi  "; System.out.println(s.trim());"#);
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn string_index_of_finds_first_occurrence() {
    let out = run_main(r#"String s = "banana"; System.out.println(s.indexOf("na"));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn string_equals_compares_content() {
    let out = run_main(
        r#"String a = "java"; String b = "java"; String c = "kotlin"; System.out.println(a.equals(b)); System.out.println(a.equals(c));"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_is_empty_on_zero_length() {
    let out = run_main(r#"String s = ""; System.out.println(s.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}
