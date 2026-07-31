use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn evaluate_true_single_when() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 5.",
        "    EVALUATE TRUE\n        WHEN N > 3\n            DISPLAY \"BIG\"\n        WHEN OTHER\n            DISPLAY \"SMALL\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["BIG"]);
}

#[test]
fn evaluate_true_when_other_triggered() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 1.",
        "    EVALUATE TRUE\n        WHEN N > 10\n            DISPLAY \"BIG\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OTHER"]);
}

#[test]
fn evaluate_subject_matching_literal() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 3.",
        "    EVALUATE N\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN 3\n            DISPLAY \"THREE\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["THREE"]);
}

#[test]
fn evaluate_range_when() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 45.",
        "    EVALUATE N\n        WHEN 1 THRU 25\n            DISPLAY \"LOW\"\n        WHEN 26 THRU 75\n            DISPLAY \"MID\"\n        WHEN 76 THRU 99\n            DISPLAY \"HIGH\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["MID"]);
}

#[test]
fn evaluate_multiple_when_same_action() {
    let out = run_prints(&p(
        "01 C PIC X VALUE \"B\".",
        "    EVALUATE C\n        WHEN \"A\"\n        WHEN \"B\"\n            DISPLAY \"A OR B\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["A OR B"]);
}

#[test]
fn evaluate_false_inverts_condition() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 3.",
        "    EVALUATE FALSE\n        WHEN N > 10\n            DISPLAY \"NOT BIG\"\n        WHEN OTHER\n            DISPLAY \"BIG\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["NOT BIG"]);
}

#[test]
fn evaluate_also_two_subjects() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    EVALUATE A ALSO B\n        WHEN 1 ALSO 2\n            DISPLAY \"ONE-TWO\"\n        WHEN OTHER ALSO OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["ONE-TWO"]);
}

#[test]
fn evaluate_also_any_wildcard() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 3.",
        "    EVALUATE A ALSO B\n        WHEN ANY ALSO 3\n            DISPLAY \"B IS 3\"\n        WHEN OTHER ALSO OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["B IS 3"]);
}

#[test]
fn evaluate_when_other_last() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 9.",
        "    EVALUATE N\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN OTHER\n            DISPLAY \"MANY\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["MANY"]);
}

#[test]
fn evaluate_thru_range_low_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 10.",
        "    EVALUATE N\n        WHEN 10 THRU 20\n            DISPLAY \"IN\"\n        WHEN OTHER\n            DISPLAY \"OUT\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["IN"]);
}

#[test]
fn evaluate_thru_range_high_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 20.",
        "    EVALUATE N\n        WHEN 10 THRU 20\n            DISPLAY \"IN\"\n        WHEN OTHER\n            DISPLAY \"OUT\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["IN"]);
}

#[test]
fn evaluate_thru_range_out_of_range() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 21.",
        "    EVALUATE N\n        WHEN 10 THRU 20\n            DISPLAY \"IN\"\n        WHEN OTHER\n            DISPLAY \"OUT\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OUT"]);
}

#[test]
fn evaluate_string_subject() {
    let out = run_prints(&p(
        "01 S PIC X(3) VALUE \"YES\".",
        "    EVALUATE S\n        WHEN \"YES\"\n            DISPLAY \"AFFIRMATIVE\"\n        WHEN \"NO\"\n            DISPLAY \"NEGATIVE\"\n        WHEN OTHER\n            DISPLAY \"UNKNOWN\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["AFFIRMATIVE"]);
}

#[test]
fn evaluate_true_compound_when_condition() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 5.",
        "    EVALUATE TRUE\n        WHEN A = 5 AND B = 5\n            DISPLAY \"BOTH FIVE\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["BOTH FIVE"]);
}

#[test]
fn evaluate_not_value_compiles() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 3.",
        "    EVALUATE N\n        WHEN NOT 1\n            DISPLAY \"NOT ONE\"\n        WHEN OTHER\n            DISPLAY \"ONE\"\n    END-EVALUATE.",
    ));
}

#[test]
fn evaluate_first_match_wins() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 5.",
        "    EVALUATE N\n        WHEN 1 THRU 10\n            DISPLAY \"FIRST\"\n        WHEN 5\n            DISPLAY \"SECOND\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["FIRST"]);
}

#[test]
fn evaluate_empty_string_matches_spaces() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE SPACES.",
        "    EVALUATE S\n        WHEN SPACES\n            DISPLAY \"BLANK\"\n        WHEN OTHER\n            DISPLAY \"NON-BLANK\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["BLANK"]);
}

#[test]
fn evaluate_nested_compiles() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 1.\n01 Y PIC 9 VALUE 2.",
        "    EVALUATE X\n        WHEN 1\n            EVALUATE Y\n                WHEN 2\n                    DISPLAY \"1-2\"\n                WHEN OTHER\n                    DISPLAY \"1-OTHER\"\n            END-EVALUATE\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
}

#[test]
fn evaluate_zero_range_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    EVALUATE N\n        WHEN 0 THRU 5\n            DISPLAY \"LOW\"\n        WHEN OTHER\n            DISPLAY \"HIGH\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["LOW"]);
}

#[test]
fn evaluate_also_any_first_subject_matches_all() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 7.\n01 B PIC 9 VALUE 7.",
        "    EVALUATE A ALSO B\n        WHEN ANY ALSO ANY\n            DISPLAY \"CATCH-ALL\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["CATCH-ALL"]);
}

#[test]
fn evaluate_grade_classification() {
    let out = run_prints(&p(
        "01 SCORE PIC 9(3) VALUE 85.",
        "    EVALUATE SCORE\n        WHEN 90 THRU 100\n            DISPLAY \"A\"\n        WHEN 80 THRU 89\n            DISPLAY \"B\"\n        WHEN 70 THRU 79\n            DISPLAY \"C\"\n        WHEN OTHER\n            DISPLAY \"F\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["B"]);
}

#[test]
fn evaluate_day_of_week_names() {
    let out = run_prints(&p(
        "01 DAY PIC 9 VALUE 1.",
        "    EVALUATE DAY\n        WHEN 1\n            DISPLAY \"MON\"\n        WHEN 2\n            DISPLAY \"TUE\"\n        WHEN 3\n            DISPLAY \"WED\"\n        WHEN 4\n            DISPLAY \"THU\"\n        WHEN 5\n            DISPLAY \"FRI\"\n        WHEN OTHER\n            DISPLAY \"WKD\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["MON"]);
}

#[test]
fn evaluate_also_mismatched_falls_through_to_other() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 9.",
        "    EVALUATE A ALSO B\n        WHEN 1 ALSO 2\n            DISPLAY \"1-2\"\n        WHEN OTHER ALSO OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OTHER"]);
}

#[test]
fn evaluate_true_or_when_condition() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 7.",
        "    EVALUATE TRUE\n        WHEN N = 5 OR N = 7\n            DISPLAY \"FIVE OR SEVEN\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["FIVE OR SEVEN"]);
}

#[test]
fn evaluate_action_adds_to_var() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 2.\n01 R PIC 9(3) VALUE 0.",
        "    EVALUATE N\n        WHEN 1\n            ADD 10 TO R\n        WHEN 2\n            ADD 20 TO R\n        WHEN OTHER\n            ADD 30 TO R\n    END-EVALUATE.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["20"]);
}

#[test]
fn evaluate_when_other_without_prior_matches() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    EVALUATE N\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN OTHER\n            DISPLAY \"ZERO\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn evaluate_compute_before_evaluate() {
    let out = run_prints(&p(
        "01 X PIC 9(2) VALUE 0.\n01 CATEGORY PIC X(6) VALUE SPACES.",
        "    COMPUTE X = 3 * 7.\n    EVALUATE X\n        WHEN 21\n            MOVE \"TWENTY\" TO CATEGORY\n        WHEN OTHER\n            MOVE \"OTHER\" TO CATEGORY\n    END-EVALUATE.\n    DISPLAY CATEGORY.",
    ));
    assert_eq!(out, vec!["TWENTY"]);
}

#[test]
fn evaluate_subject_variable_changed_mid_program() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"A\".",
        "    MOVE \"B\" TO STATUS.\n    EVALUATE STATUS\n        WHEN \"A\"\n            DISPLAY \"A\"\n        WHEN \"B\"\n            DISPLAY \"B\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["B"]);
}

#[test]
fn evaluate_true_not_condition_branch() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    EVALUATE TRUE\n        WHEN NOT N > 0\n            DISPLAY \"ZERO OR NEG\"\n        WHEN OTHER\n            DISPLAY \"POS\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["ZERO OR NEG"]);
}

#[test]
fn evaluate_numeric_zero_value_when() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    EVALUATE N\n        WHEN ZERO\n            DISPLAY \"ZERO\"\n        WHEN OTHER\n            DISPLAY \"NONZERO\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}
