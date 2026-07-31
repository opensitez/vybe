use super::helpers::compile_ok;

fn prog(data: &str, procs: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T IS RECURSIVE.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, procs
    )
}

fn prog_nonrecursive(data: &str, procs: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, procs
    )
}

#[test]
fn program_id_is_recursive_compiles() {
    compile_ok(&prog("", "    CONTINUE."));
}

#[test]
fn program_id_recursive_with_ws() {
    compile_ok(&prog(
        "01 N PIC 9(4) VALUE 0.",
        "    ADD 1 TO N.",
    ));
}

#[test]
fn program_id_recursive_with_display() {
    compile_ok(&prog(
        "01 MSG PIC X(10) VALUE \"HELLO\".",
        "    DISPLAY MSG.",
    ));
}

#[test]
fn program_id_recursive_with_loop() {
    compile_ok(&prog(
        "01 I PIC 9(3) VALUE 0.",
        "    PERFORM UNTIL I >= 10\n        ADD 1 TO I\n    END-PERFORM.",
    ));
}

#[test]
fn program_id_recursive_with_perform_para() {
    compile_ok(&prog(
        "01 R PIC 9(4) VALUE 0.",
        r#"    PERFORM CALC.
    STOP RUN.
CALC.
    ADD 42 TO R."#,
    ));
}

#[test]
fn program_id_recursive_with_if() {
    compile_ok(&prog(
        "01 X PIC 9 VALUE 5.",
        "    IF X > 3\n        DISPLAY \"BIG\"\n    ELSE\n        DISPLAY \"SMALL\"\n    END-IF.",
    ));
}

#[test]
fn program_id_recursive_with_evaluate() {
    compile_ok(&prog(
        "01 N PIC 9 VALUE 2.",
        "    EVALUATE N\n        WHEN 1 DISPLAY \"ONE\"\n        WHEN 2 DISPLAY \"TWO\"\n        WHEN OTHER DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
}

#[test]
fn program_id_recursive_with_table() {
    compile_ok(&prog(
        "01 T.\n   05 E PIC 9(2) OCCURS 5 TIMES.",
        "    MOVE 42 TO E(3).",
    ));
}

#[test]
fn program_id_recursive_with_comp_field() {
    compile_ok(&prog(
        "01 N PIC 9(8) COMP VALUE 0.",
        "    ADD 1 TO N.",
    ));
}

#[test]
fn program_id_recursive_with_signed_field() {
    compile_ok(&prog(
        "01 N PIC S9(5) VALUE -100.",
        "    ADD 200 TO N.",
    ));
}

#[test]
fn program_id_recursive_with_inspect() {
    compile_ok(&prog(
        "01 S PIC X(10) VALUE \"HELLO\".\n01 C PIC 9(3) VALUE 0.",
        "    INSPECT S TALLYING C FOR ALL \"L\".",
    ));
}

#[test]
fn program_id_recursive_with_string_op() {
    compile_ok(&prog(
        "01 A PIC X(5) VALUE \"HELLO\".\n01 R PIC X(15) VALUE SPACES.",
        "    STRING A DELIMITED BY SIZE INTO R.",
    ));
}

#[test]
fn program_id_recursive_with_unstring() {
    compile_ok(&prog(
        "01 SRC PIC X(10) VALUE \"A,B\".\n01 F1 PIC X(5).\n01 F2 PIC X(5).",
        "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2.",
    ));
}

#[test]
fn program_id_recursive_with_initialize() {
    compile_ok(&prog(
        "01 S PIC X(5) VALUE \"HELLO\".\n01 N PIC 9(5) VALUE 99999.",
        "    INITIALIZE S N.",
    ));
}

#[test]
fn program_id_recursive_with_level88() {
    compile_ok(&prog(
        "01 FLAG PIC X VALUE \"N\".\n    88 ENABLED VALUE \"Y\".",
        "    SET ENABLED TO TRUE.",
    ));
}

#[test]
fn program_id_recursive_arithmetic_sequence() {
    compile_ok(&prog(
        "01 A PIC 9(4) VALUE 10.\n01 B PIC 9(4) VALUE 20.\n01 R PIC 9(5) VALUE 0.",
        "    ADD A B GIVING R.\n    SUBTRACT A FROM R.\n    MULTIPLY 2 BY R.\n    DIVIDE 4 INTO R.",
    ));
}

#[test]
fn program_id_recursive_with_go_to() {
    compile_ok(&prog(
        "01 FLAG PIC 9 VALUE 1.",
        r#"    IF FLAG = 0
        GO TO DONE
    END-IF.
    DISPLAY "NOT DONE".
DONE.
    DISPLAY "DONE"."#,
    ));
}

#[test]
fn program_id_recursive_compute_chain() {
    compile_ok(&prog(
        "01 X PIC 9(3) VALUE 5.\n01 Y PIC 9(4) VALUE 0.",
        "    COMPUTE Y = X ** 2 + 2 * X + 1.",
    ));
}

#[test]
fn program_id_recursive_group_data() {
    compile_ok(&prog(
        "01 EMPLOYEE.\n   05 EMP-ID PIC 9(6) VALUE 100001.\n   05 EMP-NAME PIC X(20) VALUE \"ALICE\".",
        "    DISPLAY EMP-ID.\n    DISPLAY EMP-NAME.",
    ));
}

#[test]
fn program_id_recursive_with_redefines() {
    compile_ok(&prog(
        "01 UNION-FIELD PIC X(4) VALUE \"ABCD\".\n01 UNION-NUM REDEFINES UNION-FIELD PIC 9(4).",
        "    DISPLAY UNION-NUM.",
    ));
}

#[test]
fn program_id_not_recursive_by_default_compiles() {
    compile_ok(&prog_nonrecursive(
        "01 N PIC 9(4) VALUE 0.",
        "    ADD 1 TO N.",
    ));
}

#[test]
fn program_id_recursive_with_perform_n_times() {
    compile_ok(&prog(
        "01 C PIC 9(3) VALUE 0.",
        "    PERFORM 10 TIMES\n        ADD 1 TO C\n    END-PERFORM.",
    ));
}

#[test]
fn program_id_recursive_with_perform_varying() {
    compile_ok(&prog(
        "01 I PIC 9(3) VALUE 0.\n01 S PIC 9(5) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 100\n        ADD I TO S\n    END-PERFORM.",
    ));
}

#[test]
fn program_id_recursive_with_accept_from_date() {
    compile_ok(&prog(
        "01 TODAY PIC 9(6).",
        "    ACCEPT TODAY FROM DATE.",
    ));
}

#[test]
fn program_id_recursive_stop_run_in_branch() {
    compile_ok(&prog(
        "01 ERR PIC 9 VALUE 0.",
        "    IF ERR > 0\n        STOP RUN\n    END-IF.\n    DISPLAY \"OK\".",
    ));
}

#[test]
fn program_id_recursive_with_binary_fields() {
    compile_ok(&prog(
        "01 COUNT PIC 9(9) COMP VALUE 0.\n01 TOTAL PIC 9(12) COMP-3 VALUE 0.",
        "    ADD 1 TO COUNT.\n    ADD 100 TO TOTAL.",
    ));
}

#[test]
fn program_id_recursive_two_sections() {
    compile_ok(&prog(
        "01 X PIC 9 VALUE 0.",
        r#"    PERFORM SEC-A.
    PERFORM SEC-B.
    STOP RUN.
SEC-A SECTION.
    ADD 1 TO X.
SEC-B SECTION.
    ADD 2 TO X."#,
    ));
}

#[test]
fn program_id_recursive_with_nested_if() {
    compile_ok(&prog(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.",
        "    IF A > 0\n        IF B > 0\n            DISPLAY \"BOTH\"\n        END-IF\n    END-IF.",
    ));
}

#[test]
fn program_id_recursive_with_scope_terminators() {
    compile_ok(&prog(
        "01 N PIC 9(3) VALUE 0.",
        "    ADD 1 TO N\n    END-ADD.\n    SUBTRACT 1 FROM N\n    END-SUBTRACT.\n    COMPUTE N = N * 2\n    END-COMPUTE.",
    ));
}

#[test]
fn program_id_recursive_with_formatted_output() {
    compile_ok(&prog(
        "01 FORMATTED PIC ZZ9 VALUE 42.",
        "    DISPLAY FORMATTED.",
    ));
}
