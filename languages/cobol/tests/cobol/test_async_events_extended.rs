use super::helpers::{compile_ok, run_prints};

#[test]
fn call_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-A.\nPROCEDURE DIVISION.\n    CALL \"SUBA\".\n    STOP RUN.",
    );
}

#[test]
fn call_using_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-B.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 V PIC 9(3) VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUBB\" USING V.\n    STOP RUN.",
    );
}

#[test]
fn call_on_exception_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-C.\nPROCEDURE DIVISION.\n    CALL \"SUBC\"\n        ON EXCEPTION DISPLAY \"E\"\n        NOT ON EXCEPTION DISPLAY \"O\"\n    END-CALL.\n    STOP RUN.",
    );
}

#[test]
fn perform_until_with_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-D.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 I PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM UNTIL I >= 2\n        ADD 1 TO I\n        CALL \"SUBD\"\n    END-PERFORM.\n    STOP RUN.",
    );
}

#[test]
fn evaluate_with_calls_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-E.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 K PIC 9 VALUE 2.\nPROCEDURE DIVISION.\n    EVALUATE K\n        WHEN 1 CALL \"S1\"\n        WHEN 2 CALL \"S2\"\n        WHEN OTHER CALL \"SX\"\n    END-EVALUATE.\n    STOP RUN.",
    );
}

#[test]
fn call_missing_program_hits_exception() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-EX.\nPROCEDURE DIVISION.\n    CALL \"NONEXIST\"\n        ON EXCEPTION\n            DISPLAY \"ERROR\"\n        NOT ON EXCEPTION\n            DISPLAY \"SHOULD-NOT\"\n    END-CALL\n    STOP RUN.",
    );
    assert_eq!(out, vec!["ERROR"]);
}

#[test]
fn evaluate_without_target_call_syntax_still_compiles() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. C-YN.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 K PIC 9 VALUE 2.\nPROCEDURE DIVISION.\n    EVALUATE K\n        WHEN 1\n            CALL \"SUB\" ON EXCEPTION DISPLAY \"ONE\" NOT ON EXCEPTION DISPLAY \"ONE-OK\" END-CALL\n        WHEN 2\n            CALL \"SUB\" ON EXCEPTION DISPLAY \"TWO\" NOT ON EXCEPTION DISPLAY \"TWO-OK\" END-CALL\n        WHEN OTHER\n            CALL \"SUB\" ON EXCEPTION DISPLAY \"OTHER\" NOT ON EXCEPTION DISPLAY \"OTHER-OK\" END-CALL\n    END-EVALUATE\n    STOP RUN.",
    );
    assert_eq!(out, vec!["TWO"]);
}
