use crate::helpers::{run_in_main, run_main};

#[test]
fn text_block_two_line_content_preserves_internal_newline() {
    let out = run_main("String s = \"\"\"\nline1\nline2\n\"\"\"; System.out.println(s.length());");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn text_block_single_line_without_leading_indent() {
    let out = run_main("String s = \"\"\"\nhello\n\"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["hello", ""]);
}

#[test]
fn text_block_indentation_stripping_on_marked_lines() {
    let out =
        run_main("String s = \"\"\"\n    alpha\n    beta\n    \"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["alpha", "beta", ""]);
}

#[test]
fn text_block_embedded_double_quotes_preserved() {
    let out = run_main("String s = \"\"\"\n\"quoted\"\n\"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["\"quoted\"", ""]);
}

#[test]
fn text_block_embedded_single_quotes_preserved() {
    let out = run_main("String s = \"\"\"\nit's fine\n\"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["it's fine", ""]);
}

#[test]
fn text_block_empty_body_between_delimiters() {
    let out = run_main("String s = \"\"\"\n\"\"\"; System.out.println(s.length());");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn text_block_concatenation_with_regular_string() {
    let out = run_main("String s = \"pre-\" + \"\"\"\nfix\n\"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["pre-fix", ""]);
}

#[test]
fn text_block_starts_with_letter_sequence() {
    let out = run_main("String s = \"\"\"\nabc\n\"\"\"; System.out.println(s.charAt(0));");
    assert_eq!(out, vec!["a"]);
}

#[test]
fn text_block_trailing_newline_included_in_length() {
    let out = run_main("String s = \"\"\"\none\n\"\"\"; System.out.println(s.endsWith(\"\\n\"));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_multiline_joined_by_plus() {
    let out = run_main(
        "String s = \"\"\"\npart1\n\"\"\" + \"\"\"\npart2\n\"\"\"; System.out.println(s);",
    );
    assert_eq!(out, vec!["part1", "part2", ""]);
}

#[test]
fn text_block_used_as_method_argument() {
    let out = run_main("int n = \"\"\"\n123\n\"\"\".trim().length(); System.out.println(n);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn text_block_equals_same_literal_reconstruction() {
    let out = run_main(
        "String a = \"\"\"\nxy\n\"\"\"; String b = \"\"\"\nxy\n\"\"\"; System.out.println(a.equals(b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_with_spaces_only_line() {
    let out = run_main("String s = \"\"\"\n   \n\"\"\"; System.out.println(s.trim().length());");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn text_block_three_line_poem_shape() {
    let out = run_main(
        "String s = \"\"\"\n    row1\n    row2\n    row3\n    \"\"\"; System.out.println(s.split(\"\\n\").length);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn text_block_contains_backslash_character() {
    let out = run_main(
        "String s = \"\"\"\npath\\\\to\n\"\"\"; System.out.println(s.contains(\"\\\\\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_numeric_characters_parse_as_string() {
    let out =
        run_main("String s = \"\"\"\n404\n\"\"\"; System.out.println(Integer.parseInt(s.trim()));");
    assert_eq!(out, vec!["404"]);
}

#[test]
fn text_block_assign_to_final_local() {
    let out = run_main("final String s = \"\"\"\nfixed\n\"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["fixed", ""]);
}

#[test]
fn text_block_in_array_initializer() {
    let out = run_main("String[] arr = {\"a\", \"\"\"\nb\n\"\"\"}; System.out.println(arr[1]);");
    assert_eq!(out, vec!["b", ""]);
}

#[test]
fn text_block_returned_from_helper_method() {
    let types = r#"
        static String blob() {
            return """
                inner
                """;
        }
    "#;
    let out = run_in_main("System.out.println(blob());", types);
    assert_eq!(out, vec!["inner", ""]);
}

#[test]
fn text_block_indentation_stripping_keeps_relative_indent() {
    let out = run_main(
        "String s = \"\"\"\n    outer\n        inner\n    \"\"\"; System.out.println(s.startsWith(\"outer\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_line_with_embedded_spaces() {
    let out = run_main(
        "String s = \"\"\"\nspaced words\n\"\"\"; System.out.println(s.indexOf(\"words\"));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn text_block_is_not_empty_for_nonempty_body() {
    let out = run_main("String s = \"\"\"\nx\n\"\"\"; System.out.println(s.isEmpty());");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn text_block_uppercase_via_bound_call() {
    let out = run_main("String s = \"\"\"\nup\n\"\"\"; System.out.println(s.toUpperCase());");
    assert_eq!(out, vec!["UP", ""]);
}

#[test]
fn text_block_replace_internal_newline_with_dash() {
    let out =
        run_main("String s = \"\"\"\na\nb\n\"\"\"; System.out.println(s.replace(\"\\n\", \"-\"));");
    assert_eq!(out, vec!["a-b-"]);
}

#[test]
fn text_block_stored_in_field_via_constructor() {
    let types = r#"
        static class Holder {
            String text;
            Holder() { text = """
                held
                """; }
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(); System.out.println(h.text);",
        types,
    );
    assert_eq!(out, vec!["held", ""]);
}

#[test]
fn text_block_with_tab_character_inside() {
    let out =
        run_main("String s = \"\"\"\ncol\\tone\n\"\"\"; System.out.println(s.contains(\"\\t\"));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_compare_to_identical_second_block() {
    let out = run_main(
        "String a = \"\"\"\nzz\n\"\"\"; String b = \"\"\"\nzz\n\"\"\"; System.out.println(a.compareTo(b));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn text_block_substring_first_character() {
    let out = run_main("String s = \"\"\"\nmarker\n\"\"\"; System.out.println(s.substring(0, 1));");
    assert_eq!(out, vec!["m"]);
}

#[test]
fn text_block_split_into_two_lines() {
    let out = run_main(
        "String s = \"\"\"\nfirst\nsecond\n\"\"\"; System.out.println(s.split(\"\\n\").length);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn text_block_embedded_quotes_and_text_mixed() {
    let out = run_main(
        "String s = \"\"\"\nsay \"hi\" now\n\"\"\"; System.out.println(s.contains(\"hi\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn text_block_passed_to_println_directly() {
    let out = run_main("System.out.println(\"\"\"\nraw\n\"\"\");");
    assert_eq!(out, vec!["raw", ""]);
}

#[test]
fn text_block_with_leading_digit_line() {
    let out = run_main("String s = \"\"\"\n2024\n\"\"\"; System.out.println(s.trim());");
    assert_eq!(out, vec!["2024"]);
}

#[test]
fn text_block_hash_character_in_content() {
    let out = run_main("String s = \"\"\"\n#tag\n\"\"\"; System.out.println(s.charAt(0));");
    assert_eq!(out, vec!["#"]);
}

#[test]
fn text_block_multiple_adjacent_words() {
    let out = run_main(
        "String s = \"\"\"\none two three\n\"\"\"; System.out.println(s.split(\" \").length);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn text_block_indentation_stripping_on_closing_delimiter_line() {
    let out = run_main("String s = \"\"\"\n        deep\n    \"\"\"; System.out.println(s);");
    assert_eq!(out, vec!["    deep", ""]);
}

#[test]
fn text_block_repeat_twice_doubles_length_pattern() {
    let out = run_main("String s = \"\"\"\nab\n\"\"\"; System.out.println((s + s).length());");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn text_block_contains_at_sign_email_shape() {
    let out = run_main("String s = \"\"\"\nuser@host\n\"\"\"; System.out.println(s.indexOf('@'));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn text_block_starts_with_open_brace_json_shape() {
    let out =
        run_main("String s = \"\"\"\n{\n  \"k\":1\n}\n\"\"\"; System.out.println(s.charAt(0));");
    assert_eq!(out, vec!["{"]);
}

#[test]
fn text_block_assign_from_expression_chain() {
    let out = run_main("String s = (\"\"\"\nchain\n\"\"\").trim(); System.out.println(s);");
    assert_eq!(out, vec!["chain"]);
}

#[test]
fn text_block_four_line_code_snippet_line_count() {
    let out = run_main(
        "String s = \"\"\"\n    a\n    b\n    c\n    d\n    \"\"\"; System.out.println(s.lines().count());",
    );
    assert_eq!(out, vec!["4"]);
}
