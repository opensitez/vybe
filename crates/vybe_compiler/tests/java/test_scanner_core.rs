use crate::helpers::run_main;

#[test]
fn scanner_next_int_reads_single_positive_value() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("42"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn scanner_next_int_reads_first_of_several_integers() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("10 20 30"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn scanner_next_int_reads_second_token_after_first_consumed() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("10 20 30"); sc.nextInt(); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn scanner_next_int_reads_third_token_in_sequence() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("1 2 3"); sc.nextInt(); sc.nextInt(); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn scanner_next_int_parses_negative_integer() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("-17"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["-17"]);
}

#[test]
fn scanner_next_int_parses_zero() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("0"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn scanner_next_int_skips_leading_whitespace() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("   99"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn scanner_next_int_reads_large_positive_value() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("2147483647"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn scanner_next_line_reads_entire_line_without_delimiter() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("hello world"); System.out.println(sc.nextLine());"#,
    );
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn scanner_next_line_reads_text_before_newline() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("alpha\nbeta"); System.out.println(sc.nextLine());"#,
    );
    assert_eq!(out, vec!["alpha"]);
}

#[test]
fn scanner_next_line_reads_second_line_after_first() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("first\nsecond"); sc.nextLine(); System.out.println(sc.nextLine());"#,
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn scanner_next_line_on_empty_input_yields_empty_string() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner(""); System.out.println(sc.nextLine().length());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn scanner_next_returns_first_word_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("foo bar"); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["foo"]);
}

#[test]
fn scanner_next_returns_second_word_after_first() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("foo bar"); sc.next(); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["bar"]);
}

#[test]
fn scanner_next_reads_token_after_integer() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("42 items"); sc.nextInt(); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["items"]);
}

#[test]
fn scanner_next_reads_integer_as_string_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("123"); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn scanner_has_next_true_when_tokens_remain() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("a b"); System.out.println(sc.hasNext());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_false_after_all_tokens_consumed() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("solo"); sc.next(); System.out.println(sc.hasNext());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_has_next_true_after_partial_consumption() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("one two"); sc.next(); System.out.println(sc.hasNext());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_int_true_before_integer_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("55 text"); System.out.println(sc.hasNextInt());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_int_false_before_non_integer_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("hello"); System.out.println(sc.hasNextInt());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_has_next_int_false_after_last_integer_consumed() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("7"); sc.nextInt(); System.out.println(sc.hasNextInt());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_use_delimiter_comma_splits_csv_tokens() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("a,b,c"); sc.useDelimiter(","); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn scanner_use_delimiter_comma_reads_integer_fields() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("10,20,30"); sc.useDelimiter(","); System.out.println(sc.nextInt()); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn scanner_use_delimiter_semicolon_splits_fields() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("x;y;z"); sc.useDelimiter(";"); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn scanner_use_delimiter_pipe_splits_fields() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("red|green|blue"); sc.useDelimiter("\\|"); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["red", "green"]);
}

#[test]
fn scanner_use_delimiter_colon_reads_key_value_pair() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("key:value"); sc.useDelimiter(":"); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["key", "value"]);
}

#[test]
fn scanner_default_whitespace_delimiter_splits_words() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("one  two   three"); System.out.println(sc.next()); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["one", "two", "three"]);
}

#[test]
fn scanner_mixed_int_and_string_tokens_in_order() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("3 apples 2 pears"); System.out.println(sc.nextInt()); System.out.println(sc.next()); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["3", "apples", "2"]);
}

#[test]
fn scanner_next_double_reads_decimal_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("3.14"); System.out.println(sc.nextDouble());"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn scanner_next_long_reads_large_integer_token() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("9223372036854775807"); System.out.println(sc.nextLong());"#,
    );
    assert_eq!(out, vec!["9223372036854775807"]);
}

#[test]
fn scanner_next_boolean_reads_true_literal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("true"); System.out.println(sc.nextBoolean());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_next_boolean_reads_false_literal() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("false"); System.out.println(sc.nextBoolean());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_tokens_from_multiline_string() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("line1\nline2\nline3"); System.out.println(sc.next()); System.out.println(sc.next());"#,
    );
    assert_eq!(out, vec!["line1", "line2"]);
}

#[test]
fn scanner_skip_advances_past_matching_pattern() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("junk 42"); sc.skip("[^0-9]*"); System.out.println(sc.nextInt());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn scanner_find_in_line_returns_matching_substring() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("id=77 rest"); System.out.println(sc.findInLine("\\d+"));"#,
    );
    assert_eq!(out, vec!["77"]);
}

#[test]
fn scanner_has_next_line_true_when_line_remainder_exists() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("pending\n"); System.out.println(sc.hasNextLine());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn scanner_has_next_line_false_after_all_lines_read() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("only"); sc.nextLine(); System.out.println(sc.hasNextLine());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn scanner_close_does_not_block_prior_reads() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("done"); System.out.println(sc.next()); sc.close();"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn scanner_repeated_next_int_on_space_separated_run() {
    let out = run_main(
        r#"java.util.Scanner sc = new java.util.Scanner("4 8 12 16"); System.out.println(sc.nextInt() + sc.nextInt() + sc.nextInt() + sc.nextInt());"#,
    );
    assert_eq!(out, vec!["40"]);
}
