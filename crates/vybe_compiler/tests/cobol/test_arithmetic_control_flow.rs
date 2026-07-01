use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn subtract_from_updates_target() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 20.\n01 B PIC 9(3) VALUE 7.",
        "    SUBTRACT B FROM A.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["13"]);
}

#[test]
fn multiply_by_updates_target() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 6.\n01 B PIC 9(3) VALUE 7.",
        "    MULTIPLY A BY B.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn divide_into_updates_target() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 3.\n01 B PIC 9(3) VALUE 9.",
        "    DIVIDE A INTO B.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn compute_respects_parentheses() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R = (2 + 3) * 4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["20"]);
}

#[test]
fn if_else_takes_true_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 1.",
        "    IF X = 1 DISPLAY \"TRUE\" ELSE DISPLAY \"FALSE\" END-IF.",
    ));
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn evaluate_when_other_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 9.",
        "    EVALUATE X\n        WHEN 1 DISPLAY \"A\"\n        WHEN 2 DISPLAY \"B\"\n        WHEN OTHER DISPLAY \"Z\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn perform_until_counts_three_iterations() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I = 3\n        ADD 1 TO I\n    END-PERFORM.\n    DISPLAY I.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn perform_times_executes_fixed_count() {
    let out = run_prints(&p(
        "01 C PIC 9 VALUE 0.",
        "    PERFORM 4 TIMES\n        ADD 1 TO C\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["4"]);
}

#[test]
fn search_linear_with_at_end_compiles() {
    compile_ok(&p(
        "01 TBL.\n   05 E OCCURS 3 TIMES INDEXED BY IDX.\n      10 K PIC 9.\n01 F PIC X VALUE \"N\".",
        "    MOVE 1 TO K(1).\n    MOVE 2 TO K(2).\n    MOVE 3 TO K(3).\n    SET IDX TO 1.\n    SEARCH E\n        AT END MOVE \"N\" TO F\n        WHEN K(IDX) = 2 MOVE \"Y\" TO F\n    END-SEARCH.\n    DISPLAY F.",
    ));
}

#[test]
fn call_on_exception_not_on_exception_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUB-X\"\n        ON EXCEPTION DISPLAY \"ERR\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn goback_statement_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GOBACK.");
}

#[test]
fn cancel_literal_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CANCEL \"SUBMOD\".\n    STOP RUN.",
    );
}

#[test]
fn set_index_up_and_down_compiles() {
    compile_ok(&p(
        "01 TAB PIC 9 OCCURS 5 TIMES INDEXED BY I.",
        "    SET I TO 1.\n    SET I UP BY 2.\n    SET I DOWN BY 1.",
    ));
}

#[test]
fn condition_name_level_88_in_if() {
    let out = run_prints(&p(
        "01 WS-S PIC X VALUE \"A\".\n   88 IS-A VALUE \"A\".\n   88 IS-B VALUE \"B\".",
        "    IF IS-A\n        DISPLAY \"A-STATE\"\n    ELSE\n        DISPLAY \"OTHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["A-STATE"]);
}

#[test]
fn initialize_group_sets_children() {
    let out = run_prints(&p(
        "01 G.\n   05 A PIC 9 VALUE 5.\n   05 B PIC X VALUE \"Z\".",
        "    INITIALIZE G.\n    DISPLAY A.\n    DISPLAY B.",
    ));
    assert_eq!(out, vec!["5", "Z"]);
}

#[test]
fn inspect_tallying_specific_character_counts_total() {
    let out = run_prints(&p(
        "01 TXT PIC X(8) VALUE \"ABABXABA\".\n01 CNT PIC 9(2) VALUE 0.",
        "    INSPECT TXT TALLYING CNT FOR ALL \"A\".\n    DISPLAY CNT.",
    ));
    assert_eq!(out, vec!["4"]);
}

#[test]
fn unstring_with_count_in_compiles() {
    compile_ok(&p(
        "01 SRC PIC X(10) VALUE \"AA,BBB\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).\n01 C1 PIC 9(2).\n01 C2 PIC 9(2).",
        "    UNSTRING SRC DELIMITED BY \",\"\n        INTO F1 COUNT IN C1\n             F2 COUNT IN C2.",
    ));
}

#[test]
fn inspect_replacing_first_compiles() {
    compile_ok(&p(
        "01 TXT PIC X(6) VALUE \"AAAAAA\".",
        "    INSPECT TXT REPLACING FIRST \"A\" BY \"B\".",
    ));
}
