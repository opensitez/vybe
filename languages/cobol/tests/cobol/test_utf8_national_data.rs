use super::helpers::compile_ok;

#[test]
fn utf8_national_literal_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    MOVE N\"CAFE\" TO N1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_display_of_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION DISPLAY-OF(N1) TO D1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_national_of_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION NATIONAL-OF(D1) TO N1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_second_literal_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    MOVE N\"HELLO\" TO N1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_move_between_national_items_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 N2 PIC N(10).\nPROCEDURE DIVISION.\n    MOVE N1 TO N2.\n    STOP RUN.",
    );
}

#[test]
fn utf8_initialize_national_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    INITIALIZE N1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_compare_national_literal_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\nPROCEDURE DIVISION.\n    IF N1 = N\"CAFE\" DISPLAY \"Y\" END-IF.\n    STOP RUN.",
    );
}

#[test]
fn utf8_display_conversion_roundtrip_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF8.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION DISPLAY-OF(N1) TO D1.\n    MOVE FUNCTION NATIONAL-OF(D1) TO N1.\n    STOP RUN.",
    );
}

#[test]
fn utf8_national_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 NG.\n   05 N1 PIC N(5).\n   05 N2 PIC N(5).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn utf8_function_pair_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. UTF10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 N1 PIC N(10).\n01 D1 PIC X(10).\nPROCEDURE DIVISION.\n    MOVE FUNCTION NATIONAL-OF(D1) TO N1.\n    MOVE FUNCTION DISPLAY-OF(N1) TO D1.\n    STOP RUN.",
    );
}
