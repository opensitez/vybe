use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn special_return_code_move_zero() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    MOVE 0 TO RETURN-CODE.",
    ));
}

#[test]
fn special_return_code_move_nonzero() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    MOVE 8 TO RETURN-CODE.",
    ));
}

#[test]
fn special_return_code_display() {
    let out = run_prints(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    MOVE 0 TO RETURN-CODE.\n    DISPLAY RETURN-CODE.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn special_return_code_used_in_if() {
    let out = run_prints(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    IF RETURN-CODE = 0\n        DISPLAY \"SUCCESS\"\n    ELSE\n        DISPLAY \"FAILURE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["SUCCESS"]);
}

#[test]
fn tally_special_register_compiles() {
    compile_ok(&p("", "    MOVE 0 TO TALLY."));
}

#[test]
fn tally_after_inspect_compiles() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"AABABAB\".",
        "    INSPECT S TALLYING TALLY FOR ALL \"A\".",
    ));
}

#[test]
fn tally_initial_value_zero_compiles() {
    compile_ok(&p("", "    DISPLAY TALLY."));
}

#[test]
fn address_of_ws_field_compiles() {
    compile_ok(&p(
        "01 N PIC 9(4) VALUE 0.\n01 PTR USAGE POINTER.",
        "    SET PTR TO ADDRESS OF N.",
    ));
}

#[test]
fn address_of_group_item_compiles() {
    compile_ok(&p(
        "01 GRP.\n   05 F1 PIC X(5) VALUE \"HELLO\".\n01 PTR USAGE POINTER.",
        "    SET PTR TO ADDRESS OF GRP.",
    ));
}

#[test]
fn address_of_in_condition_compiles() {
    compile_ok(&p(
        "01 N PIC 9(4) VALUE 0.\n01 PTR USAGE POINTER.",
        "    SET PTR TO ADDRESS OF N.\n    IF PTR NOT = NULL\n        DISPLAY \"HAS ADDR\"\n    END-IF.",
    ));
}

#[test]
fn pointer_null_comparison_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER VALUE NULL.",
        "    IF P = NULL\n        DISPLAY \"NULL\"\n    END-IF.",
    ));
}

#[test]
fn pointer_set_to_null() {
    compile_ok(&p("01 P USAGE POINTER.", "    SET P TO NULL."));
}

#[test]
fn special_register_in_if_and_compute() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    COMPUTE RETURN-CODE = 4 + 4.",
    ));
}

#[test]
fn return_code_set_on_error_path() {
    let out = run_prints(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.\n01 ERR PIC 9 VALUE 1.",
        "    IF ERR > 0\n        MOVE 4 TO RETURN-CODE\n    END-IF.\n    DISPLAY RETURN-CODE.",
    ));
    assert_eq!(out, vec!["0004"]);
}

#[test]
fn return_code_success_path() {
    let out = run_prints(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.\n01 ERR PIC 9 VALUE 0.",
        "    IF ERR > 0\n        MOVE 8 TO RETURN-CODE\n    ELSE\n        MOVE 0 TO RETURN-CODE\n    END-IF.\n    DISPLAY RETURN-CODE.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn address_of_elementary_then_display_ws() {
    compile_ok(&p(
        "01 DATA-ITEM PIC X(10) VALUE \"HELLO\".\n01 DATA-PTR USAGE POINTER.",
        "    SET DATA-PTR TO ADDRESS OF DATA-ITEM.\n    DISPLAY DATA-ITEM.",
    ));
}

#[test]
fn special_register_length_of_compiles() {
    compile_ok(&p(
        "01 S PIC X(10) VALUE \"HELLO\".\n01 L PIC 9(5) VALUE 0.",
        "    COMPUTE L = FUNCTION LENGTH(S).",
    ));
}

#[test]
fn special_lineage_counter_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT REPORT-FILE ASSIGN TO \"output.txt\".\nDATA DIVISION.\nFILE SECTION.\nFD REPORT-FILE LINAGE IS 60 LINES.\n01 REPORT-REC PIC X(80).\nWORKING-STORAGE SECTION.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn return_code_in_evaluate() {
    let out = run_prints(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    MOVE 4 TO RETURN-CODE.\n    EVALUATE RETURN-CODE\n        WHEN 0 DISPLAY \"OK\"\n        WHEN 4 DISPLAY \"WARN\"\n        WHEN 8 DISPLAY \"ERR\"\n        WHEN OTHER DISPLAY \"UNKNOWN\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["WARN"]);
}

#[test]
fn pointer_assign_and_compare_two_pointers() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 P1 USAGE POINTER.\n01 P2 USAGE POINTER.",
        "    SET P1 TO ADDRESS OF A.\n    SET P2 TO ADDRESS OF A.\n    IF P1 = P2\n        DISPLAY \"SAME\"\n    END-IF.",
    ));
}

#[test]
fn tally_used_as_counter_then_displayed() {
    compile_ok(&p(
        "01 S PIC X(15) VALUE \"MISSISSIPPI    \".",
        "    MOVE 0 TO TALLY.\n    INSPECT S TALLYING TALLY FOR ALL \"S\".\n    DISPLAY TALLY.",
    ));
}

#[test]
fn return_code_can_be_used_in_add() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    ADD 4 TO RETURN-CODE.",
    ));
}

#[test]
fn address_of_table_element_compiles() {
    compile_ok(&p(
        "01 T.\n   05 E PIC X(5) OCCURS 10 TIMES.\n01 P USAGE POINTER.",
        "    SET P TO ADDRESS OF E(1).",
    ));
}

#[test]
fn special_registers_all_used_in_one_program() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.\n01 S PIC X(10) VALUE \"AABABAB\".\n01 N PIC 9(5) VALUE 0.\n01 P USAGE POINTER.",
        "    MOVE 0 TO RETURN-CODE.\n    MOVE 0 TO TALLY.\n    INSPECT S TALLYING TALLY FOR ALL \"A\".\n    SET P TO ADDRESS OF N.\n    IF RETURN-CODE = 0 AND TALLY > 0 AND P NOT = NULL\n        DISPLAY \"ALL VALID\"\n    END-IF.",
    ));
}

#[test]
fn return_code_subtract() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 8.",
        "    SUBTRACT 4 FROM RETURN-CODE.",
    ));
}

#[test]
fn address_of_in_loop_compiles() {
    compile_ok(&p(
        "01 ITEM PIC X(10) VALUE \"DATA\".\n01 P USAGE POINTER.\n01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I >= 3\n        ADD 1 TO I\n        SET P TO ADDRESS OF ITEM\n    END-PERFORM.",
    ));
}

#[test]
fn return_code_conditional_stop_run() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4) VALUE 0.",
        "    IF RETURN-CODE > 0\n        STOP RUN\n    END-IF.\n    DISPLAY \"CONTINUED\".",
    ));
}
