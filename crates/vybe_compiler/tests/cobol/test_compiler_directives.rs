use super::helpers::compile_ok;

#[test]
fn eject_directive_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. DIR1.\nEJECT\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn title_directive_compiles() {
    compile_ok(
        "TITLE \"COBOL TEST\"\nIDENTIFICATION DIVISION.\nPROGRAM-ID. DIR2.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn skip1_directive_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. DIR3.\nSKIP1\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn skip2_directive_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. DIR4.\nSKIP2\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn skip3_directive_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. DIR5.\nSKIP3\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn page_directive_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. DIR6.\nPAGE\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn title_and_eject_directives_compiles() {
    compile_ok(
        "TITLE \"DIR\"\nIDENTIFICATION DIVISION.\nPROGRAM-ID. DIR7.\nEJECT\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_free_directive_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. DIR8.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_fixed_directive_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FIXED\nIDENTIFICATION DIVISION.\nPROGRAM-ID. DIR9.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn multiple_listing_directives_compiles() {
    compile_ok(
        "TITLE \"LISTING\"\nSKIP1\nIDENTIFICATION DIVISION.\nPROGRAM-ID. DIR10.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}
