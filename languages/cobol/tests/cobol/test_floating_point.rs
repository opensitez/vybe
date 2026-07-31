use super::helpers::{compile_ok, run_prints};

#[test]
fn comp1_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 F1 USAGE COMP-1.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn comp2_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 F2 USAGE COMP-2.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn floating_compute_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-1.\n01 B USAGE COMP-2.\n01 C USAGE COMP-2.\nPROCEDURE DIVISION.\n    COMPUTE C = A + B.\n    STOP RUN.",
    );
}

#[test]
fn comp1_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-1.\n01 B USAGE COMP-1.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}

#[test]
fn comp2_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-2.\n01 B USAGE COMP-2.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}

#[test]
fn floating_add_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-1.\n01 B USAGE COMP-1.\nPROCEDURE DIVISION.\n    ADD B TO A.\n    STOP RUN.",
    );
}

#[test]
fn floating_subtract_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-2.\n01 B USAGE COMP-2.\nPROCEDURE DIVISION.\n    SUBTRACT B FROM A.\n    STOP RUN.",
    );
}

#[test]
fn floating_multiply_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP8.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-1.\n01 B USAGE COMP-1.\nPROCEDURE DIVISION.\n    MULTIPLY A BY B.\n    STOP RUN.",
    );
}

#[test]
fn floating_divide_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-2.\n01 B USAGE COMP-2.\nPROCEDURE DIVISION.\n    DIVIDE A INTO B.\n    STOP RUN.",
    );
}

#[test]
fn floating_compute_parentheses_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-2.\n01 B USAGE COMP-2.\n01 C USAGE COMP-2.\nPROCEDURE DIVISION.\n    COMPUTE C = (A + B) * B.\n    STOP RUN.",
    );
}

#[test]
fn floating_add_runtime_outputs_expected_value() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. FP11.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A USAGE COMP-1 VALUE 1.\n01 B USAGE COMP-1 VALUE 2.\n01 C USAGE COMP-1.\nPROCEDURE DIVISION.\n    ADD B TO A\n    MOVE A TO C\n    DISPLAY C\n    STOP RUN.",
    );
    assert_eq!(out, vec!["3"]);
}
