use super::helpers::{compile_ok, run_prints};

#[test]
fn call_no_args_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"M1\".\n    STOP RUN.",
    );
}
#[test]
fn call_one_arg_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC X(5).\nPROCEDURE DIVISION.\n    CALL \"M2\" USING A.\n    STOP RUN.",
    );
}
#[test]
fn call_two_args_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC X(5).\n01 B PIC 9(3).\nPROCEDURE DIVISION.\n    CALL \"M3\" USING A B.\n    STOP RUN.",
    );
}
#[test]
fn call_by_reference_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(5).\nPROCEDURE DIVISION.\n    CALL \"M4\" USING BY REFERENCE A.\n    STOP RUN.",
    );
}
#[test]
fn call_by_content_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(5).\nPROCEDURE DIVISION.\n    CALL \"M5\" USING BY CONTENT A.\n    STOP RUN.",
    );
}
#[test]
fn call_by_value_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(5).\nPROCEDURE DIVISION.\n    CALL \"M6\" USING BY VALUE A.\n    STOP RUN.",
    );
}
#[test]
fn call_on_exception_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"M7\" ON EXCEPTION DISPLAY \"E\" END-CALL.\n    STOP RUN.",
    );
}
#[test]
fn call_not_on_exception_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"M8\" NOT ON EXCEPTION DISPLAY \"OK\" END-CALL.\n    STOP RUN.",
    );
}
#[test]
fn cancel_module_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CANCEL \"M9\".\n    STOP RUN.",
    );
}
#[test]
fn chained_calls_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"A\".\n    CALL \"B\".\n    STOP RUN.",
    );
}
#[test]
fn call_after_if_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    IF X = 1 CALL \"C\" END-IF.\n    STOP RUN.",
    );
}
#[test]
fn call_in_perform_loop_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM 2 TIMES\n        CALL \"D\"\n    END-PERFORM.\n    STOP RUN.",
    );
}
#[test]
fn nested_program_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. OUTER.\nPROCEDURE DIVISION.\n    STOP RUN.\nEND PROGRAM OUTER.",
    );
}
#[test]
fn call_returning_handle_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 H PIC X(20).\nPROCEDURE DIVISION.\n    CALL \"E\" RETURNING H.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_call_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N PIC X(10) VALUE \"M10\".\nPROCEDURE DIVISION.\n    CALL N.\n    STOP RUN.",
    );
}
#[test]
fn module_sequence_a_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"M11\".\n    CALL \"M12\".\n    CALL \"M13\".\n    STOP RUN.",
    );
}
#[test]
fn module_sequence_b_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"INIT\".\n    CALL \"RUN\".\n    CALL \"DONE\".\n    STOP RUN.",
    );
}
#[test]
fn call_with_group_arg_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 G.\n   05 A PIC X(3).\n   05 B PIC 9(2).\nPROCEDURE DIVISION.\n    CALL \"MG\" USING G.\n    STOP RUN.",
    );
}

#[test]
fn call_exception_runtime_message() {
    let output = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"M7\"\n        ON EXCEPTION\n            DISPLAY \"E\"\n        NOT ON EXCEPTION\n            DISPLAY \"OK\"\n    END-CALL\n    STOP RUN.",
    );
    assert_eq!(output, vec!["E"]);
}

#[test]
fn call_dynamic_name_runtime() {
    let output = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N PIC X(10) VALUE \"NO-MOD\".\nPROCEDURE DIVISION.\n    CALL N\n        ON EXCEPTION\n            DISPLAY \"MISS\"\n    END-CALL\n    STOP RUN.",
    );
    assert_eq!(output, vec!["MISS"]);
}

#[test]
fn call_by_reference_runtime_value_preserved() {
    let output = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(5) VALUE 12345.\nPROCEDURE DIVISION.\n    CALL \"M4\" USING BY REFERENCE A\n        ON EXCEPTION\n            DISPLAY A\n    END-CALL\n    STOP RUN.",
    );
    assert_eq!(output, vec!["12345"]);
}
