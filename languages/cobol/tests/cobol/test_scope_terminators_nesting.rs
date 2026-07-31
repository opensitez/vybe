use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn scope_end_if_closes_branch() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 5.",
        "    IF N > 0\n        DISPLAY \"POS\"\n    END-IF.\n    DISPLAY \"AFTER\".",
    ));
    assert_eq!(out, vec!["POS", "AFTER"]);
}

#[test]
fn scope_end_perform_closes_loop() {
    let out = run_prints(&p(
        "",
        "    PERFORM 2 TIMES\n        DISPLAY \"LOOP\"\n    END-PERFORM.\n    DISPLAY \"DONE\".",
    ));
    assert_eq!(out, vec!["LOOP", "LOOP", "DONE"]);
}

#[test]
fn scope_end_evaluate_compiles() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 1.",
        "    EVALUATE N\n        WHEN 1\n            DISPLAY \"ONE\"\n    END-EVALUATE.",
    ));
}

#[test]
fn scope_end_add_not_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 0.",
        "    ADD 1 TO A\n    END-ADD.",
    ));
}

#[test]
fn scope_end_subtract_not_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 100.",
        "    SUBTRACT 1 FROM A\n    END-SUBTRACT.",
    ));
}

#[test]
fn scope_end_multiply_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 2.",
        "    MULTIPLY 3 BY A\n    END-MULTIPLY.",
    ));
}

#[test]
fn scope_end_divide_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 10.\n01 R PIC 9(3) VALUE 0.",
        "    DIVIDE 2 INTO A GIVING R\n    END-DIVIDE.",
    ));
}

#[test]
fn scope_end_compute_compiles() {
    compile_ok(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 2 + 2\n    END-COMPUTE.",
    ));
}

#[test]
fn scope_end_if_else_both_branches() {
    let out = run_prints(&p(
        "01 F PIC 9 VALUE 0.",
        "    IF F = 0\n        DISPLAY \"ZERO\"\n    ELSE\n        DISPLAY \"NONZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn scope_nested_end_if_inside_perform() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 S PIC 9(2) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        IF I > 3\n            ADD I TO S\n        END-IF\n    END-PERFORM.\n    DISPLAY S.",
    ));
    // 4+5 = 9
    assert_eq!(out, vec!["09"]);
}

#[test]
fn scope_end_string_compiles() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 B PIC X(5) VALUE \"WORLD\".\n01 R PIC X(20).",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R\n    END-STRING.",
    ));
}

#[test]
fn scope_end_unstring_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(10) VALUE \"A,B,C\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2\n    END-UNSTRING.",
    ));
}

#[test]
fn scope_nested_end_evaluate_inside_if() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    IF A > 0\n        EVALUATE B\n            WHEN 2\n                DISPLAY \"A-POS-B-2\"\n        END-EVALUATE\n    END-IF.",
    ));
}

#[test]
fn scope_end_perform_with_until() {
    let out = run_prints(&p(
        "01 C PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL C >= 5\n        ADD 1 TO C\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn scope_end_if_prevents_dangling_else() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 1.\n01 Y PIC 9 VALUE 0.",
        "    IF X = 1\n        IF Y = 1\n            DISPLAY \"INNER\"\n        END-IF\n    ELSE\n        DISPLAY \"OUTER-ELSE\"\n    END-IF.\n    DISPLAY \"AFTER\".",
    ));
    assert_eq!(out, vec!["AFTER"]);
}

#[test]
fn scope_end_add_with_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 999.",
        "    ADD 1 TO A\n    ON SIZE ERROR\n        DISPLAY \"OVERFLOW\"\n    END-ADD.",
    ));
}

#[test]
fn scope_end_subtract_with_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 0.",
        "    SUBTRACT 1 FROM A\n    ON SIZE ERROR\n        DISPLAY \"UNDERFLOW\"\n    END-SUBTRACT.",
    ));
}

#[test]
fn scope_end_compute_with_on_size_error() {
    compile_ok(&p(
        "01 R PIC 9(2) VALUE 0.",
        "    COMPUTE R = 99 * 99\n    ON SIZE ERROR\n        DISPLAY \"TOO BIG\"\n    END-COMPUTE.",
    ));
}

#[test]
fn scope_three_consecutive_end_if() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.\n01 C PIC 9 VALUE 3.",
        "    IF A = 1\n        DISPLAY \"A\"\n    END-IF.\n    IF B = 2\n        DISPLAY \"B\"\n    END-IF.\n    IF C = 3\n        DISPLAY \"C\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn scope_end_if_inside_evaluate() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 1.",
        "    EVALUATE N\n        WHEN 1\n            IF N = 1\n                DISPLAY \"ONE\"\n            END-IF\n    END-EVALUATE.",
    ));
}

#[test]
fn scope_end_multiply_with_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 999.",
        "    MULTIPLY 999 BY A\n    ON SIZE ERROR\n        DISPLAY \"OVERFLOW\"\n    END-MULTIPLY.",
    ));
}

#[test]
fn scope_end_divide_with_remainder() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 17.\n01 Q PIC 9(3) VALUE 0.\n01 REM PIC 9(3) VALUE 0.",
        "    DIVIDE 5 INTO A GIVING Q REMAINDER REM\n    END-DIVIDE.",
    ));
}

#[test]
fn scope_end_search_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X OCCURS 5 TIMES.",
        "    SEARCH E\n        AT END\n            DISPLAY \"NOT FOUND\"\n        WHEN E(1) = \"A\"\n            DISPLAY \"FOUND\"\n    END-SEARCH.",
    ));
}

#[test]
fn scope_end_perform_varying_nested_end_if() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 OUT PIC X VALUE \"N\".",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        IF I = 3\n            MOVE \"Y\" TO OUT\n        END-IF\n    END-PERFORM.\n    DISPLAY OUT.",
    ));
    assert_eq!(out, vec!["Y"]);
}

#[test]
fn scope_end_if_then_perform() {
    let out = run_prints(&p(
        "01 FLAG PIC X VALUE \"Y\".",
        "    IF FLAG = \"Y\"\n        PERFORM 2 TIMES\n            DISPLAY \"RUN\"\n        END-PERFORM\n    END-IF.",
    ));
    assert_eq!(out, vec!["RUN", "RUN"]);
}

#[test]
fn scope_end_evaluate_runs_when_other() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 7.",
        "    EVALUATE N\n        WHEN 1\n            DISPLAY \"ONE\"\n        WHEN OTHER\n            DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OTHER"]);
}

#[test]
fn scope_end_add_not_size_error_branch() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 5.",
        "    ADD 10 TO A\n    NOT ON SIZE ERROR\n        DISPLAY \"OK\"\n    END-ADD.",
    ));
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn scope_double_nested_end_perform() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 J PIC 9 VALUE 0.\n01 S PIC 9(3) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2\n            ADD 1 TO S\n        END-PERFORM\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["6"]);
}
