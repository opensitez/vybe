use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn inspect_converting_single_char_pair() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"HELLO\".",
        "    INSPECT S CONVERTING \"L\" TO \"R\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["HERRO     "]);
}

#[test]
fn inspect_converting_multi_char_mapping() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"aeiou\".",
        "    INSPECT S CONVERTING \"aeiou\" TO \"AEIOU\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["AEIOU"]);
}

#[test]
fn inspect_converting_no_match_leaves_unchanged() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    INSPECT S CONVERTING \"xyz\" TO \"XYZ\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn inspect_tallying_all_occurrence_count() {
    let out = run_prints(&p(
        "01 S PIC X(15) VALUE \"ABRACADABRA\".\n01 C PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"A\".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn inspect_tallying_leading_zeros() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"000123\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR LEADING \"0\".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn inspect_tallying_characters_in_alphanumeric() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"MISSISSIPPI\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"S\".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["4"]);
}

#[test]
fn inspect_replacing_all_with_space() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"HELLO\".",
        "    INSPECT S REPLACING ALL \"L\" BY \" \".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["HE  O     "]);
}

#[test]
fn inspect_replacing_leading_zeros_with_spaces() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"000042\".",
        "    INSPECT S REPLACING LEADING \"0\" BY \" \".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["   042  "]);
}

#[test]
fn inspect_replacing_first_char() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"AABABC\".",
        "    INSPECT S REPLACING FIRST \"A\" BY \"X\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["XABABC    "]);
}

#[test]
fn inspect_tallying_before_delimiter() {
    compile_ok(&p(
        "01 S PIC X(20) VALUE \"HELLO WORLD\".\n01 C PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\" BEFORE \" \".",
    ));
}

#[test]
fn inspect_tallying_after_delimiter() {
    compile_ok(&p(
        "01 S PIC X(20) VALUE \"HELLO WORLD\".\n01 C PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\" AFTER \" \".",
    ));
}

#[test]
fn inspect_converting_digit_to_asterisk() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"A1B2C3D4\".",
        "    INSPECT S CONVERTING \"1234567890\" TO \"**********\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["A*B*C*D*"]);
}

#[test]
fn inspect_tallying_zero_when_not_found() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"ABCDE\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"Z\".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn inspect_replacing_all_preserves_non_matching() {
    let out = run_prints(&p(
        "01 S PIC X(7) VALUE \"ABCABCA\".",
        "    INSPECT S REPLACING ALL \"A\" BY \"X\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["XBCXBCX"]);
}

#[test]
fn inspect_tallying_and_replacing_combined() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO\".\n01 C PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\" REPLACING ALL \"L\" BY \"R\".",
    ));
}

#[test]
fn inspect_converting_lowercase_to_uppercase_partial() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"abcXYZ\".",
        "    INSPECT S CONVERTING \"abcdefghij\" TO \"ABCDEFGHIJ\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["ABCXYZ  "]);
}

#[test]
fn inspect_replacing_trailing_chars_compiles() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO     \".",
        "    INSPECT S REPLACING TRAILING \" \" BY \"_\".",
    ));
}

#[test]
fn inspect_tallying_length_difference() {
    // Count spaces in padded string
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"AB\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \" \".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["8"]);
}

#[test]
fn inspect_converting_entire_string_single_char() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"AAAAA\".",
        "    INSPECT S CONVERTING \"A\" TO \"B\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["BBBBB"]);
}

#[test]
fn inspect_replacing_all_with_zero() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE \"X1X2X3\".",
        "    INSPECT S REPLACING ALL \"X\" BY \"0\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["010203"]);
}

#[test]
fn inspect_tallying_for_characters() {
    let out = run_prints(&p(
        "01 S PIC X(26) VALUE \"ABCDEFGHIJKLMNOPQRSTUVWXYZ\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR CHARACTERS.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["26"]);
}

#[test]
fn inspect_replacing_characters_with_star() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO\".",
        "    INSPECT S REPLACING CHARACTERS BY \"*\".",
    ));
}

#[test]
fn inspect_tallying_multiple_targets() {
    compile_ok(&p(
        "01 S PIC X(20) VALUE \"HELLO WORLD\".\n01 C1 PIC 9(3) VALUE 0.\n01 C2 PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C1 FOR ALL \"L\" C2 FOR ALL \"O\".",
    ));
}

#[test]
fn inspect_converting_digit_sequence() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"1234567890\".",
        "    INSPECT S CONVERTING \"0123456789\" TO \"9876543210\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["8765432109"]);
}

#[test]
fn inspect_replacing_first_space() {
    let out = run_prints(&p(
        "01 S PIC X(10) VALUE \"A B C\".",
        "    INSPECT S REPLACING FIRST \" \" BY \"_\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["A_B C     "]);
}

#[test]
fn inspect_tallying_before_first_space() {
    let out = run_prints(&p(
        "01 S PIC X(15) VALUE \"HELLO WORLD\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\" BEFORE INITIAL \" \".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn inspect_tallying_after_initial_delimiter() {
    let out = run_prints(&p(
        "01 S PIC X(15) VALUE \"HELLO WORLD\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\" AFTER INITIAL \" \".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn inspect_replacing_leading_zeros_runtime() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"00000042\".",
        "    INSPECT S REPLACING LEADING \"0\" BY \" \".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["      42"]);
}

#[test]
fn inspect_converting_space_to_underscore() {
    let out = run_prints(&p(
        "01 S PIC X(12) VALUE \"HELLO WORLD\".",
        "    INSPECT S CONVERTING \" \" TO \"_\".\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["HELLO_WORLD "]);
}

#[test]
fn inspect_tallying_characters_after_replacing() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE \"AABBCC\".\n01 C PIC 9(2) VALUE 0.",
        "    INSPECT S REPLACING ALL \"A\" BY \"X\".\n    INSPECT S TALLYING C FOR ALL \"X\".\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["2"]);
}
