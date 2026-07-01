use super::helpers::compile_ok;

#[test]
fn call_external_program_no_args_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-A.\nPROCEDURE DIVISION.\n    CALL \"SUB-A\".\n    STOP RUN.",
    );
}

#[test]
fn call_external_program_with_using_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-B.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(10).\nPROCEDURE DIVISION.\n    CALL \"SUB-B\" USING WS-A.\n    STOP RUN.",
    );
}

#[test]
fn call_external_program_by_reference_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-C.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5) VALUE 10.\nPROCEDURE DIVISION.\n    CALL \"SUB-C\" USING BY REFERENCE WS-A.\n    STOP RUN.",
    );
}

#[test]
fn call_external_program_by_content_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-D.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5) VALUE 10.\nPROCEDURE DIVISION.\n    CALL \"SUB-D\" USING BY CONTENT WS-A.\n    STOP RUN.",
    );
}

#[test]
fn call_external_program_on_exception_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN-E.\nPROCEDURE DIVISION.\n    CALL \"SUB-E\"\n        ON EXCEPTION DISPLAY \"FAIL\"\n        NOT ON EXCEPTION DISPLAY \"OK\"\n    END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn nested_program_and_call_pattern_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. OUTER-PROG.\nPROCEDURE DIVISION.\n    CALL \"INNER-PROG\".\n    STOP RUN.\nEND PROGRAM OUTER-PROG.",
    );
}
