use super::helpers::compile_ok;

#[test]
fn comp_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 B1 PIC S9(4) COMP.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn comp4_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 B1 PIC S9(4) COMP-4.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn comp5_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 B1 PIC S9(9) COMP-5.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn binary_addition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 10.\n01 B PIC S9(4) COMP VALUE 5.\nPROCEDURE DIVISION.\n    ADD B TO A.\n    STOP RUN.",
    );
}

#[test]
fn binary_subtraction_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 10.\n01 B PIC S9(4) COMP VALUE 5.\nPROCEDURE DIVISION.\n    SUBTRACT B FROM A.\n    STOP RUN.",
    );
}

#[test]
fn binary_multiplication_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 3.\n01 B PIC S9(4) COMP VALUE 4.\nPROCEDURE DIVISION.\n    MULTIPLY A BY B.\n    STOP RUN.",
    );
}

#[test]
fn binary_division_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 2.\n01 B PIC S9(4) COMP VALUE 8.\nPROCEDURE DIVISION.\n    DIVIDE A INTO B.\n    STOP RUN.",
    );
}

#[test]
fn binary_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN8.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 10.\n01 B PIC S9(4) COMP.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}

#[test]
fn binary_compute_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(4) COMP VALUE 10.\n01 B PIC S9(4) COMP VALUE 5.\n01 C PIC S9(4) COMP.\nPROCEDURE DIVISION.\n    COMPUTE C = A + B.\n    STOP RUN.",
    );
}

#[test]
fn binary_comp5_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. BN10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(9) COMP-5 VALUE 99.\n01 B PIC S9(9) COMP-5.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}
