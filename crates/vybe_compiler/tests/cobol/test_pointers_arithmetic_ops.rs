use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn pointer_decl_compiles() {
    compile_ok(&p("01 P USAGE POINTER.", "    SET P TO NULL."));
}
#[test]
fn pointer_address_of_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 A PIC X(5).",
        "    SET P TO ADDRESS OF A.",
    ));
}
#[test]
fn pointer_pass_to_call_compiles() {
    compile_ok(&p("01 P USAGE POINTER.", "    CALL \"PTR-USE\" USING P."));
}
#[test]
fn pointer_compare_null_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.",
        "    IF P = NULL DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn pointer_assign_twice_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 A PIC X(5).\n01 B PIC X(5).",
        "    SET P TO ADDRESS OF A.\n    SET P TO ADDRESS OF B.",
    ));
}
#[test]
fn pointer_function_pointer_compiles() {
    compile_ok(&p("01 FP USAGE FUNCTION-POINTER.", "    DISPLAY \"FP\"."));
}
#[test]
fn pointer_procedure_pointer_compiles() {
    compile_ok(&p("01 PP USAGE PROCEDURE-POINTER.", "    DISPLAY \"PP\"."));
}
#[test]
fn pointer_callback_call_compiles() {
    compile_ok(&p(
        "01 PP USAGE PROCEDURE-POINTER.",
        "    CALL \"INVOKE-CB\" USING PP.",
    ));
}
#[test]
fn pointer_store_in_table_compiles() {
    compile_ok(&p(
        "01 PT OCCURS 3 TIMES USAGE POINTER.\n01 P USAGE POINTER.",
        "    SET P TO NULL.",
    ));
}
#[test]
fn pointer_loop_assign_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 I PIC 9 VALUE 1.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        SET P TO NULL\n    END-PERFORM.",
    ));
}
#[test]
fn pointer_arith_external_call_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 N PIC 9(3) VALUE 1.",
        "    CALL \"PTR-ADD\" USING P N.",
    ));
}
#[test]
fn pointer_diff_external_call_compiles() {
    compile_ok(&p(
        "01 P1 USAGE POINTER.\n01 P2 USAGE POINTER.\n01 D PIC 9(5).",
        "    CALL \"PTR-DIFF\" USING P1 P2 D.",
    ));
}
#[test]
fn pointer_increment_external_call_compiles() {
    compile_ok(&p("01 P USAGE POINTER.", "    CALL \"PTR-INC\" USING P."));
}
#[test]
fn pointer_decrement_external_call_compiles() {
    compile_ok(&p("01 P USAGE POINTER.", "    CALL \"PTR-DEC\" USING P."));
}
#[test]
fn pointer_cast_external_call_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 N PIC 9(10).",
        "    CALL \"PTR-TO-NUM\" USING P N.",
    ));
}
#[test]
fn pointer_from_num_external_call_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.\n01 N PIC 9(10) VALUE 100.",
        "    CALL \"NUM-TO-PTR\" USING N P.",
    ));
}
#[test]
fn pointer_safety_check_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER.",
        "    IF P NOT = NULL DISPLAY \"OK\" END-IF.",
    ));
}
#[test]
fn pointer_reset_null_compiles() {
    compile_ok(&p("01 P USAGE POINTER.", "    SET P TO NULL."));
}
