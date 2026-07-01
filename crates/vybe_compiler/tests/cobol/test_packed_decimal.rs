use super::helpers::compile_ok;

#[test]
fn comp3_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 P1 PIC S9(5)V99 COMP-3.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_add_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 10.\n01 B PIC S9(3) COMP-3 VALUE 5.\nPROCEDURE DIVISION.\n    ADD B TO A.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_redefines_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 P1 PIC S9(3) COMP-3 VALUE 123.\n01 P1-X REDEFINES P1 PIC X(2).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_subtract_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 10.\n01 B PIC S9(3) COMP-3 VALUE 2.\nPROCEDURE DIVISION.\n    SUBTRACT B FROM A.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_multiply_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 3.\n01 B PIC S9(3) COMP-3 VALUE 4.\nPROCEDURE DIVISION.\n    MULTIPLY A BY B.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_divide_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 2.\n01 B PIC S9(3) COMP-3 VALUE 8.\nPROCEDURE DIVISION.\n    DIVIDE A INTO B.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 123.\n01 B PIC S9(3) COMP-3.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_compute_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD8.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 10.\n01 B PIC S9(3) COMP-3 VALUE 20.\n01 C PIC S9(4) COMP-3.\nPROCEDURE DIVISION.\n    COMPUTE C = A + B.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_with_signed_fraction_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3)V9 COMP-3 VALUE 12.3.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn packed_decimal_with_edited_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. PD10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC S9(3) COMP-3 VALUE 123.\n01 B PIC ZZZ.\nPROCEDURE DIVISION.\n    MOVE A TO B.\n    STOP RUN.",
    );
}
