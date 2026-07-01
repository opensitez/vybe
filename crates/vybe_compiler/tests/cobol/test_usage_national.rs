use super::helpers::compile_ok;

#[test]
fn usage_national_definition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 USAGE NATIONAL PIC N(10).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn national_literal_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    MOVE N\"HELLO\" TO N1.\n    STOP RUN.",
    );
}

#[test]
fn national_comparison_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(5).\nPROCEDURE DIVISION.\n    IF N1 = N\"A\" DISPLAY \"Y\" END-IF.\n    STOP RUN.",
    );
}

#[test]
fn national_group_item_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 NGROUP.\n   05 N1 PIC N(5).\n   05 N2 PIC N(5).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn national_move_between_items_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(5).\n01 N2 PIC N(5).\nPROCEDURE DIVISION.\n    MOVE N1 TO N2.\n    STOP RUN.",
    );
}

#[test]
fn national_initialize_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    INITIALIZE N1.\n    STOP RUN.",
    );
}

#[test]
fn national_string_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    MOVE N\"WORLD\" TO N1.\n    STOP RUN.",
    );
}

#[test]
fn national_display_of_function_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT8.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION DISPLAY-OF(N1) TO D1.\n    STOP RUN.",
    );
}

#[test]
fn national_of_function_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION NATIONAL-OF(D1) TO N1.\n    STOP RUN.",
    );
}

#[test]
fn national_if_comparison_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. NAT10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(5).\nPROCEDURE DIVISION.\n    IF N1 = N\"HELLO\" DISPLAY \"Y\" END-IF.\n    STOP RUN.",
    );
}
