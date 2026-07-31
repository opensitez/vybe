use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn raise_exception_literal_compiles() {
    compile_ok(&p("", "    RAISE EXCEPTION \"ERR\"."));
}
#[test]
fn raise_exception_then_display_compiles() {
    compile_ok(&p(
        "",
        "    RAISE EXCEPTION \"E1\".\n    DISPLAY \"AFTER\".",
    ));
}
#[test]
fn call_on_exception_compiles() {
    compile_ok(&p(
        "",
        "    CALL \"MAYBE\"\n        ON EXCEPTION DISPLAY \"FAIL\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.",
    ));
}
#[test]
fn read_at_end_branch_compiles() {
    compile_ok(&p(
        "01 F PIC X(80).",
        "    READ WS-FILE\n        AT END DISPLAY \"EOF\"\n    END-READ.",
    ));
}
#[test]
fn write_invalid_key_branch_compiles() {
    compile_ok(&p(
        "01 F PIC X(80).",
        "    WRITE F\n        INVALID KEY DISPLAY \"BAD\"\n    END-WRITE.",
    ));
}
#[test]
fn divide_size_error_branch_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 0.",
        "    DIVIDE A BY B\n        ON SIZE ERROR DISPLAY \"SE\"\n    END-DIVIDE.",
    ));
}
#[test]
fn compute_size_error_branch_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 9.\n01 B PIC 9 VALUE 9.\n01 C PIC 9 VALUE 0.",
        "    COMPUTE C = A ** B\n        ON SIZE ERROR DISPLAY \"SE\"\n    END-COMPUTE.",
    ));
}
#[test]
fn string_overflow_branch_compiles() {
    compile_ok(&p(
        "01 A PIC X(5) VALUE \"ABCDE\".\n01 B PIC X(5) VALUE \"FGHIJ\".\n01 O PIC X(5).",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O\n        ON OVERFLOW DISPLAY \"OV\"\n    END-STRING.",
    ));
}
#[test]
fn unstring_overflow_branch_compiles() {
    compile_ok(&p(
        "01 S PIC X(3) VALUE \"A,B\".\n01 A PIC X(1).",
        "    UNSTRING S DELIMITED BY \",\" INTO A\n        ON OVERFLOW DISPLAY \"OV\"\n    END-UNSTRING.",
    ));
}
#[test]
fn xml_exception_branch_compiles() {
    compile_ok(&p(
        "01 X PIC X(50).\n01 R PIC X(5).",
        "    XML GENERATE X FROM R\n        ON EXCEPTION DISPLAY \"XERR\"\n    END-XML.",
    ));
}
#[test]
fn json_exception_branch_compiles() {
    compile_ok(&p(
        "01 J PIC X(50).\n01 R PIC X(10).",
        "    JSON PARSE J INTO R.",
    ));
}
#[test]
fn sql_error_check_compiles() {
    compile_ok(&p(
        "01 SQLCODE PIC S9(9) VALUE 0.",
        "    EXEC SQL SELECT 1 END-EXEC.\n    IF SQLCODE NOT = 0 DISPLAY \"SQLERR\" END-IF.",
    ));
}
#[test]
fn rollback_on_error_compiles() {
    compile_ok(&p(
        "01 SQLCODE PIC S9(9) VALUE 1.",
        "    IF SQLCODE NOT = 0\n        EXEC SQL ROLLBACK END-EXEC\n    END-IF.",
    ));
}
#[test]
fn raise_custom_error_code_compiles() {
    compile_ok(&p(
        "01 E PIC 9(4) VALUE 1001.",
        "    CALL \"RAISE-CODE\" USING E.",
    ));
}
#[test]
fn fallback_path_with_evaluate_compiles() {
    compile_ok(&p(
        "01 ST PIC 9 VALUE 3.",
        "    EVALUATE ST\n        WHEN 1 DISPLAY \"OK\"\n        WHEN 2 DISPLAY \"WARN\"\n        WHEN OTHER DISPLAY \"ERR\"\n    END-EVALUATE.",
    ));
}
#[test]
fn retry_loop_pattern_compiles() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 0.",
        "    PERFORM UNTIL N >= 3\n        ADD 1 TO N\n        CALL \"TRY-STEP\"\n    END-PERFORM.",
    ));
}
#[test]
fn cancel_on_error_compiles() {
    compile_ok(&p(
        "",
        "    CALL \"MAYBE\"\n        ON EXCEPTION CANCEL \"MAYBE\"\n    END-CALL.",
    ));
}
#[test]
fn error_flag_set_compiles() {
    compile_ok(&p(
        "01 ERR-FLAG PIC 9 VALUE 0.",
        "    CALL \"MAYBE\"\n        ON EXCEPTION MOVE 1 TO ERR-FLAG\n    END-CALL.",
    ));
}

#[test]
fn add_on_size_error_compiles() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 999.\n01 B PIC 9(3) VALUE 2.\n01 C PIC 9(3).",
        "    ADD A TO B GIVING C\n        ON SIZE ERROR DISPLAY \"SE\"\n        NOT ON SIZE ERROR DISPLAY \"OK\"\n    END-ADD.",
    ));
}

#[test]
fn subtract_on_size_error_compiles() {
    compile_ok(&p(
        "01 A PIC 9(2) VALUE 10.\n01 B PIC 9(2) VALUE 20.\n01 C PIC 9(2).",
        "    SUBTRACT B FROM A GIVING C\n        ON SIZE ERROR DISPLAY \"SE\"\n    END-SUBTRACT.",
    ));
}

#[test]
fn multiply_size_error_compiles() {
    compile_ok(&p(
        "01 A PIC 9(4) VALUE 200.\n01 B PIC 9(4) VALUE 300.\n01 C PIC 9(4).",
        "    MULTIPLY A BY B GIVING C\n        ON SIZE ERROR DISPLAY \"MUL-SE\"\n    END-MULTIPLY.",
    ));
}

#[test]
fn string_overflow_branch_fires() {
    let out = run_prints(&p(
        "01 A PIC X(5) VALUE \"ABCDE\".\n01 B PIC X(5) VALUE \"FGHIJ\".\n01 O PIC X(5).",
        "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O\n        ON OVERFLOW DISPLAY \"OV\"\n    END-STRING.",
    ));
    assert_eq!(out, vec!["OV"]);
}

#[test]
fn unstring_overflow_branch_fires() {
    let out = run_prints(&p(
        "01 S PIC X(3) VALUE \"A,B\".\n01 A PIC X(1).",
        "    UNSTRING S DELIMITED BY \",\" INTO A\n        ON OVERFLOW DISPLAY \"OV\"\n    END-UNSTRING.",
    ));
    assert_eq!(out, vec!["OV"]);
}
