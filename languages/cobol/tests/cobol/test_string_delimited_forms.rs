use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

// ── STRING with SIZE, SPACE, and char delimiters ─────────────────
#[test]
fn string_size_delimiter_appends_full() {
    let out = run_prints(&p(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING A DELIMITED BY SIZE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO          "]);
}

#[test]
fn string_space_delimiter_stops_at_first_space() {
    let out = run_prints(&p(
        "01 A PIC X(10) VALUE \"HELLO     \".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING A DELIMITED BY SPACE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO          "]);
}

#[test]
fn string_char_delimiter_stops_before_char() {
    let out = run_prints(&p(
        "01 A PIC X(10) VALUE \"HELLO:END\".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING A DELIMITED BY \":\" INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO          "]);
}

#[test]
fn string_two_operands_size_delimiter() {
    let out = run_prints(&p(
        "01 A PIC X(3) VALUE \"ABC\".\n01 B PIC X(3) VALUE \"DEF\".\n01 R PIC X(10) VALUE SPACES.",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["ABCDEF    "]);
}

#[test]
fn string_literal_and_variable_mixed() {
    let out = run_prints(&p(
        "01 NAME PIC X(8) VALUE \"ALICE   \".\n01 R PIC X(20) VALUE SPACES.",
        "    STRING \"HELLO \" DELIMITED BY SIZE NAME DELIMITED BY SPACE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["HELLO ALICE          "]);
}

#[test]
fn string_pointer_starts_mid_buffer() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 R PIC X(20) VALUE SPACES.\n01 PTR PIC 9(3) VALUE 6.",
        "    STRING A DELIMITED BY SIZE INTO R WITH POINTER PTR.",
    ));
}

#[test]
fn string_three_vars_space_delimited() {
    let out = run_prints(&p(
        "01 A PIC X(10) VALUE \"FIRST     \".\n01 B PIC X(10) VALUE \"SECOND    \".\n01 C PIC X(10) VALUE \"THIRD     \".\n01 R PIC X(30) VALUE SPACES.",
        "    STRING A DELIMITED BY SPACE \" \" DELIMITED BY SIZE B DELIMITED BY SPACE \" \" DELIMITED BY SIZE C DELIMITED BY SPACE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["FIRST SECOND THIRD           "]);
}

#[test]
fn string_all_literal_operands() {
    let out = run_prints(&p(
        "01 R PIC X(20) VALUE SPACES.",
        "    STRING \"FOO\" DELIMITED BY SIZE \"BAR\" DELIMITED BY SIZE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["FOOBAR              "]);
}

#[test]
fn string_overflow_handler_compiles() {
    compile_ok(&p(
        "01 A PIC X(20) VALUE \"LONG VALUE HERE      \".\n01 R PIC X(5) VALUE SPACES.",
        "    STRING A DELIMITED BY SIZE INTO R\n    ON OVERFLOW\n        DISPLAY \"OVERFLOW\"\n    END-STRING.",
    ));
}

#[test]
fn string_not_overflow_compiles() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 R PIC X(10) VALUE SPACES.",
        "    STRING A DELIMITED BY SIZE INTO R\n    NOT ON OVERFLOW\n        DISPLAY \"OK\"\n    END-STRING.",
    ));
}

// ── UNSTRING with COUNT, TALLYING, multiple delimiters ──────────
#[test]
fn unstring_basic_comma_delimiter() {
    let out = run_prints(&p(
        "01 SRC PIC X(15) VALUE \"A,B,C\".\n01 F1 PIC X(5) VALUE SPACES.\n01 F2 PIC X(5) VALUE SPACES.\n01 F3 PIC X(5) VALUE SPACES.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3.\n    DISPLAY F1.\n    DISPLAY F2.\n    DISPLAY F3.",
    ));
    assert_eq!(out, vec!["A    ", "B    ", "C    "]);
}

#[test]
fn unstring_space_delimiter() {
    let out = run_prints(&p(
        "01 SRC PIC X(15) VALUE \"HELLO WORLD\".\n01 W1 PIC X(8) VALUE SPACES.\n01 W2 PIC X(8) VALUE SPACES.",
        "    UNSTRING SRC DELIMITED BY SPACE INTO W1 W2.\n    DISPLAY W1.\n    DISPLAY W2.",
    ));
    assert_eq!(out, vec!["HELLO   ", "WORLD   "]);
}

#[test]
fn unstring_with_tallying_in() {
    compile_ok(&p(
        "01 SRC PIC X(20) VALUE \"A,B,C,D\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).\n01 F3 PIC X(5).\n01 CNT PIC 9(3) VALUE 0.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3 TALLYING IN CNT.",
    ));
}

#[test]
fn unstring_with_count_in() {
    compile_ok(&p(
        "01 SRC PIC X(15) VALUE \"ABC,DEF\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).\n01 C1 PIC 9(3) VALUE 0.\n01 C2 PIC 9(3) VALUE 0.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 COUNT IN C1 F2 COUNT IN C2.",
    ));
}

#[test]
fn unstring_multiple_delimiters_or() {
    compile_ok(&p(
        "01 SRC PIC X(20) VALUE \"A,B;C D\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).\n01 F3 PIC X(5).\n01 F4 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \",\" OR \";\" OR SPACE INTO F1 F2 F3 F4.",
    ));
}

#[test]
fn unstring_all_delimiter_collapses() {
    compile_ok(&p(
        "01 SRC PIC X(20) VALUE \"A,,B\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY ALL \",\" INTO F1 F2.",
    ));
}

#[test]
fn unstring_pointer_form() {
    compile_ok(&p(
        "01 SRC PIC X(20) VALUE \"HELLO WORLD\".\n01 F1 PIC X(8).\n01 PTR PIC 9(3) VALUE 1.",
        "    UNSTRING SRC DELIMITED BY SPACE INTO F1 WITH POINTER PTR.",
    ));
}

#[test]
fn unstring_overflow_handler_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(10) VALUE \"A,B,C,D,E\".\n01 F1 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1\n    ON OVERFLOW\n        DISPLAY \"OVERFLOW\"\n    END-UNSTRING.",
    ));
}

#[test]
fn string_result_then_displayed() {
    let out = run_prints(&p(
        "01 FIRST PIC X(5) VALUE \"JOHN \".\n01 LAST PIC X(5) VALUE \"DOE  \".\n01 FULL PIC X(12) VALUE SPACES.",
        "    STRING FIRST DELIMITED BY SPACE \" \" DELIMITED BY SIZE LAST DELIMITED BY SPACE INTO FULL.\n    DISPLAY FULL.",
    ));
    assert_eq!(out, vec!["JOHN DOE    "]);
}

#[test]
fn unstring_two_fields_captured() {
    let out = run_prints(&p(
        "01 SRC PIC X(10) VALUE \"YES:NO\".\n01 F1 PIC X(5) VALUE SPACES.\n01 F2 PIC X(5) VALUE SPACES.",
        "    UNSTRING SRC DELIMITED BY \":\" INTO F1 F2.\n    DISPLAY F1.\n    DISPLAY F2.",
    ));
    assert_eq!(out, vec!["YES  ", "NO   "]);
}

#[test]
fn string_pointer_updated_compiles() {
    compile_ok(&p(
        "01 A PIC X(3) VALUE \"ABC\".\n01 B PIC X(3) VALUE \"DEF\".\n01 R PIC X(20) VALUE SPACES.\n01 PTR PIC 9(3) VALUE 1.",
        "    STRING A DELIMITED BY SIZE INTO R WITH POINTER PTR.\n    STRING B DELIMITED BY SIZE INTO R WITH POINTER PTR.",
    ));
}

#[test]
fn unstring_into_numeric_field() {
    compile_ok(&p(
        "01 SRC PIC X(10) VALUE \"123,456\".\n01 F1 PIC 9(5) VALUE 0.\n01 F2 PIC 9(5) VALUE 0.",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2.",
    ));
}

#[test]
fn string_with_five_literals() {
    let out = run_prints(&p(
        "01 R PIC X(20) VALUE SPACES.",
        "    STRING \"A\" DELIMITED BY SIZE \"B\" DELIMITED BY SIZE \"C\" DELIMITED BY SIZE \"D\" DELIMITED BY SIZE \"E\" DELIMITED BY SIZE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["ABCDE               "]);
}

#[test]
fn unstring_single_char_delim_many_fields() {
    let out = run_prints(&p(
        "01 SRC PIC X(9) VALUE \"1|2|3|4|5\".\n01 F1 PIC X.\n01 F2 PIC X.\n01 F3 PIC X.\n01 F4 PIC X.\n01 F5 PIC X.",
        "    UNSTRING SRC DELIMITED BY \"|\" INTO F1 F2 F3 F4 F5.\n    DISPLAY F1.\n    DISPLAY F2.\n    DISPLAY F5.",
    ));
    assert_eq!(out, vec!["1", "2", "5"]);
}

#[test]
fn string_space_delim_excludes_trailing_spaces() {
    let out = run_prints(&p(
        "01 A PIC X(10) VALUE \"COBOL     \".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING A DELIMITED BY SPACE \"!\" DELIMITED BY SIZE INTO R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["COBOL!         "]);
}

#[test]
fn unstring_delimiter_in_field_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(30) VALUE \"a=1;b=2\".\n01 K1 PIC X(5).\n01 V1 PIC X(5).\n01 K2 PIC X(5).\n01 V2 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \"=\" OR \";\" INTO K1 V1 K2 V2.",
    ));
}

#[test]
fn string_with_delimiters_then_inspect() {
    compile_ok(&p(
        "01 NAME PIC X(10) VALUE \"JOHN      \".\n01 R PIC X(20) VALUE SPACES.\n01 C PIC 9(2) VALUE 0.",
        "    STRING NAME DELIMITED BY SPACE INTO R.\n    INSPECT R TALLYING C FOR ALL \"O\".",
    ));
}

#[test]
fn unstring_not_overflow_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(10) VALUE \"A,B\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2\n    NOT ON OVERFLOW\n        DISPLAY \"OK\"\n    END-UNSTRING.",
    ));
}
