use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn nested_if_basic_true_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 5.\n01 Y PIC 9 VALUE 3.",
        "    IF X > 0\n        IF Y > 0\n            DISPLAY \"BOTH POS\"\n        ELSE\n            DISPLAY \"X POS Y NOT\"\n        END-IF\n    ELSE\n        DISPLAY \"X NOT POS\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BOTH POS"]);
}

#[test]
fn nested_if_inner_false_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 5.\n01 Y PIC 9 VALUE 0.",
        "    IF X > 0\n        IF Y > 0\n            DISPLAY \"BOTH POS\"\n        ELSE\n            DISPLAY \"X POS Y NOT\"\n        END-IF\n    ELSE\n        DISPLAY \"X NOT POS\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["X POS Y NOT"]);
}

#[test]
fn nested_if_outer_false_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 0.\n01 Y PIC 9 VALUE 5.",
        "    IF X > 0\n        IF Y > 0\n            DISPLAY \"BOTH POS\"\n        ELSE\n            DISPLAY \"X POS Y NOT\"\n        END-IF\n    ELSE\n        DISPLAY \"X NOT POS\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["X NOT POS"]);
}

#[test]
fn nested_if_three_levels_deep() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.\n01 C PIC 9 VALUE 3.",
        "    IF A > 0\n        IF B > 0\n            IF C > 0\n                DISPLAY \"DEEP TRUE\"\n            ELSE\n                DISPLAY \"C FAIL\"\n            END-IF\n        ELSE\n            DISPLAY \"B FAIL\"\n        END-IF\n    ELSE\n        DISPLAY \"A FAIL\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["DEEP TRUE"]);
}

#[test]
fn if_else_if_chain_first_match() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 10.",
        "    IF N < 10\n        DISPLAY \"SMALL\"\n    ELSE IF N < 20\n        DISPLAY \"MEDIUM\"\n    ELSE IF N < 100\n        DISPLAY \"LARGE\"\n    ELSE\n        DISPLAY \"HUGE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["MEDIUM"]);
}

#[test]
fn if_else_if_chain_last_match() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 200.",
        "    IF N < 10\n        DISPLAY \"SMALL\"\n    ELSE IF N < 50\n        DISPLAY \"MEDIUM\"\n    ELSE IF N < 100\n        DISPLAY \"LARGE\"\n    ELSE\n        DISPLAY \"HUGE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["HUGE"]);
}

#[test]
fn if_compound_condition_and() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 5.",
        "    IF A = 5 AND B = 5\n        DISPLAY \"BOTH FIVE\"\n    ELSE\n        DISPLAY \"NOT BOTH\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BOTH FIVE"]);
}

#[test]
fn if_compound_condition_or() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 5.",
        "    IF A = 5 OR B = 5\n        DISPLAY \"ONE FIVE\"\n    ELSE\n        DISPLAY \"NONE FIVE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ONE FIVE"]);
}

#[test]
fn nested_if_with_perform() {
    let out = run_prints(&p(
        "01 FLAG PIC 9 VALUE 1.\n01 C PIC 9 VALUE 0.",
        "    IF FLAG = 1\n        PERFORM 3 TIMES\n            ADD 1 TO C\n        END-PERFORM\n    END-IF.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn nested_if_compute_in_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 3.\n01 R PIC 9(3) VALUE 0.",
        "    IF X > 0\n        COMPUTE R = X * X\n    ELSE\n        COMPUTE R = 0 - X\n    END-IF.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn if_else_nested_set_two_vars() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 0.\n01 OUT1 PIC X(3) VALUE \"---\".\n01 OUT2 PIC X(3) VALUE \"---\".",
        "    IF A = 1\n        MOVE \"YES\" TO OUT1\n        IF B = 0\n            MOVE \"NO\" TO OUT2\n        END-IF\n    END-IF.\n    DISPLAY OUT1.\n    DISPLAY OUT2.",
    ));
    assert_eq!(out, vec!["YES", "NO"]);
}

#[test]
fn if_nested_in_else() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    IF N > 5\n        DISPLAY \"HIGH\"\n    ELSE\n        IF N > 0\n            DISPLAY \"LOW\"\n        ELSE\n            DISPLAY \"ZERO\"\n        END-IF\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn if_string_equality_nested() {
    let out = run_prints(&p(
        "01 CODE PIC X(2) VALUE \"OK\".\n01 TYPE PIC X VALUE \"A\".",
        "    IF CODE = \"OK\"\n        IF TYPE = \"A\"\n            DISPLAY \"OK-A\"\n        ELSE\n            DISPLAY \"OK-OTHER\"\n        END-IF\n    ELSE\n        DISPLAY \"NOT-OK\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["OK-A"]);
}

#[test]
fn if_nested_in_loop_conditionally_breaks_simulated() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 0.\n01 FOUND PIC X VALUE \"N\".",
        "    PERFORM UNTIL I >= 10 OR FOUND = \"Y\"\n        ADD 1 TO I\n        IF I = 5\n            MOVE \"Y\" TO FOUND\n        END-IF\n    END-PERFORM.\n    DISPLAY I.",
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn if_four_branches_chained() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 3.",
        "    IF N = 1\n        DISPLAY \"ONE\"\n    ELSE IF N = 2\n        DISPLAY \"TWO\"\n    ELSE IF N = 3\n        DISPLAY \"THREE\"\n    ELSE IF N = 4\n        DISPLAY \"FOUR\"\n    ELSE\n        DISPLAY \"OTHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["THREE"]);
}

#[test]
fn if_not_equal_moves_alternative() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".\n01 R PIC X(5) VALUE SPACES.",
        "    IF S NOT = \"HELLO\"\n        MOVE \"WRONG\" TO R\n    ELSE\n        MOVE \"RIGHT\" TO R\n    END-IF.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["RIGHT"]);
}

#[test]
fn if_greater_equal_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 10.",
        "    IF N >= 10\n        DISPLAY \"GE\"\n    ELSE\n        DISPLAY \"LT\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["GE"]);
}

#[test]
fn if_less_equal_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 10.",
        "    IF N <= 10\n        DISPLAY \"LE\"\n    ELSE\n        DISPLAY \"GT\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["LE"]);
}

#[test]
fn if_nested_display_order() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 1.",
        "    DISPLAY \"BEFORE\".\n    IF X = 1\n        DISPLAY \"INSIDE\"\n    END-IF.\n    DISPLAY \"AFTER\".",
    ));
    assert_eq!(out, vec!["BEFORE", "INSIDE", "AFTER"]);
}

#[test]
fn if_compute_in_outer_then_nested_check() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 6.\n01 B PIC 9(3) VALUE 7.\n01 PROD PIC 9(5) VALUE 0.",
        "    COMPUTE PROD = A * B.\n    IF PROD > 40\n        IF PROD < 50\n            DISPLAY \"RANGE\"\n        ELSE\n            DISPLAY \"HIGH\"\n        END-IF\n    ELSE\n        DISPLAY \"LOW\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["RANGE"]);
}

#[test]
fn if_with_continue_in_branch() {
    compile_ok(&p(
        "01 X PIC 9 VALUE 1.",
        "    IF X = 1\n        CONTINUE\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
}

#[test]
fn if_not_alphabetic_class() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"123AB\".",
        "    IF NOT S IS ALPHABETIC\n        DISPLAY \"MIXED\"\n    ELSE\n        DISPLAY \"ALPHA\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["MIXED"]);
}

#[test]
fn if_nested_three_levels_only_first_enters() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 1.\n01 C PIC 9 VALUE 1.",
        "    IF A > 0\n        IF B > 0\n            IF C > 0\n                DISPLAY \"ALL\"\n            END-IF\n        END-IF\n    ELSE\n        DISPLAY \"OUTER FALSE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["OUTER FALSE"]);
}

#[test]
fn if_equal_spaces_check() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"     \".",
        "    IF S = SPACES\n        DISPLAY \"BLANK\"\n    ELSE\n        DISPLAY \"NOT BLANK\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BLANK"]);
}

#[test]
fn if_equal_zeros_check() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE 0.",
        "    IF N = ZEROS\n        DISPLAY \"ALL ZERO\"\n    ELSE\n        DISPLAY \"NONZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ALL ZERO"]);
}

#[test]
fn if_else_moves_to_field_in_both_branches() {
    let out = run_prints(&p(
        "01 COND PIC 9 VALUE 0.\n01 RESULT PIC X(4) VALUE \"----\".",
        "    IF COND = 1\n        MOVE \"TRUE\" TO RESULT\n    ELSE\n        MOVE \"FALS\" TO RESULT\n    END-IF.\n    DISPLAY RESULT.",
    ));
    assert_eq!(out, vec!["FALS"]);
}

#[test]
fn if_with_perform_in_branch() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 1.\n01 S PIC 9(2) VALUE 0.",
        "    IF N > 0\n        PERFORM 5 TIMES\n            ADD 1 TO S\n        END-PERFORM\n    END-IF.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn if_chained_else_if_five_branches() {
    let out = run_prints(&p(
        "01 MONTH PIC 9(2) VALUE 9.",
        "    IF MONTH = 1\n        DISPLAY \"JAN\"\n    ELSE IF MONTH = 2\n        DISPLAY \"FEB\"\n    ELSE IF MONTH = 6\n        DISPLAY \"JUN\"\n    ELSE IF MONTH = 9\n        DISPLAY \"SEP\"\n    ELSE\n        DISPLAY \"OTHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["SEP"]);
}

#[test]
fn if_boundary_exclusive_gt() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 10.",
        "    IF N > 10\n        DISPLAY \"GT\"\n    ELSE\n        DISPLAY \"LE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["LE"]);
}
