use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn if_greater_than_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 10.\n01 WS-B PIC 9(3) VALUE 5.",
        "    IF WS-A > WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_less_than_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 3.\n01 WS-B PIC 9(3) VALUE 5.",
        "    IF WS-A < WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_equal_to_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 5.\n01 WS-B PIC 9(3) VALUE 5.",
        "    IF WS-A = WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_not_equal_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 5.\n01 WS-B PIC 9(3) VALUE 6.",
        "    IF WS-A NOT = WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_greater_than_or_equal_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 10.\n01 WS-B PIC 9(3) VALUE 10.",
        "    IF WS-A >= WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_less_than_or_equal_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 5.\n01 WS-B PIC 9(3) VALUE 10.",
        "    IF WS-A <= WS-B\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_and_condition_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 5.\n01 WS-B PIC 9(3) VALUE 7.",
        "    IF WS-A > 0 AND WS-B > 0\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_or_condition_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 0.\n01 WS-B PIC 9(3) VALUE 7.",
        "    IF WS-A = 0 OR WS-B = 0\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_not_condition_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 0.",
        "    IF NOT WS-A > 0\n        DISPLAY \"A\"\n    END-IF.",
    ));
}
#[test]
fn if_nested_condition_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 5.",
        "    IF WS-A > 0\n        IF WS-A < 10\n            DISPLAY \"A\"\n        END-IF\n    END-IF.",
    ));
}
#[test]
fn if_with_else_branch_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 3.",
        "    IF WS-A > 5\n        DISPLAY \"BIG\"\n    ELSE\n        DISPLAY \"SMALL\"\n    END-IF.",
    ));
}
#[test]
fn if_with_elseif_style_branch_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 2.",
        "    IF WS-A = 1\n        DISPLAY \"ONE\"\n    ELSE\n        IF WS-A = 2\n            DISPLAY \"TWO\"\n        END-IF\n    END-IF.",
    ));
}
#[test]
fn if_numeric_positive_compiles() {
    compile_ok(&p(
        "01 WS-A PIC S9(3) VALUE 5.",
        "    IF WS-A IS POSITIVE\n        DISPLAY \"POS\"\n    END-IF.",
    ));
}
#[test]
fn if_numeric_negative_compiles() {
    compile_ok(&p(
        "01 WS-A PIC S9(3) VALUE -5.",
        "    IF WS-A IS NEGATIVE\n        DISPLAY \"NEG\"\n    END-IF.",
    ));
}
#[test]
fn if_numeric_zero_compiles() {
    compile_ok(&p(
        "01 WS-A PIC S9(3) VALUE 0.",
        "    IF WS-A IS ZERO\n        DISPLAY \"ZERO\"\n    END-IF.",
    ));
}
#[test]
fn if_alphabetic_compiles() {
    compile_ok(&p(
        "01 WS-A PIC X(3) VALUE \"ABC\".",
        "    IF WS-A IS ALPHABETIC\n        DISPLAY \"ALPHA\"\n    END-IF.",
    ));
}
#[test]
fn if_alphanumeric_compiles() {
    compile_ok(&p(
        "01 WS-A PIC X(3) VALUE \"A1B\".",
        "    IF WS-A IS ALPHANUMERIC\n        DISPLAY \"ALNUM\"\n    END-IF.",
    ));
}
#[test]
fn if_numeric_class_compiles() {
    compile_ok(&p(
        "01 WS-A PIC X(3) VALUE \"123\".",
        "    IF WS-A IS NUMERIC\n        DISPLAY \"NUM\"\n    END-IF.",
    ));
}
#[test]
fn evaluate_simple_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(1) VALUE 2.",
        "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
}
#[test]
fn evaluate_true_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(2) VALUE 85.",
        "    EVALUATE TRUE\n        WHEN WS-A >= 90\n            DISPLAY \"A\"\n        WHEN WS-A >= 80\n            DISPLAY \"B\"\n        WHEN OTHER\n            DISPLAY \"F\"\n    END-EVALUATE.",
    ));
}
#[test]
fn evaluate_string_compiles() {
    compile_ok(&p(
        "01 WS-A PIC X(1) VALUE \"B\".",
        "    EVALUATE WS-A\n        WHEN \"A\"\n            DISPLAY \"ALPHA\"\n        WHEN \"B\"\n            DISPLAY \"BETA\"\n    END-EVALUATE.",
    ));
}
#[test]
fn evaluate_multiple_when_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(1) VALUE 3.",
        "    EVALUATE WS-A\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN 2\n            DISPLAY \"TWO\"\n        WHEN 3\n            DISPLAY \"THREE\"\n    END-EVALUATE.",
    ));
}
#[test]
fn perform_until_condition_compiles() {
    compile_ok(&p(
        "01 WS-I PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL WS-I >= 3\n        ADD 1 TO WS-I\n    END-PERFORM.",
    ));
}
#[test]
fn perform_varying_condition_compiles() {
    compile_ok(&p(
        "01 WS-I PIC 9(2) VALUE 0.",
        "    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5\n        DISPLAY WS-I\n    END-PERFORM.",
    ));
}
#[test]
fn perform_times_condition_compiles() {
    compile_ok(&p(
        "",
        "    PERFORM 4 TIMES\n        DISPLAY \"LOOP\"\n    END-PERFORM.",
    ));
}
#[test]
fn perform_inline_condition_compiles() {
    compile_ok(&p(
        "",
        "    PERFORM 2 TIMES\n        CONTINUE\n    END-PERFORM.",
    ));
}
