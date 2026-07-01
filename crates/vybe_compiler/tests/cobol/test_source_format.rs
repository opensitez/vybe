use super::helpers::compile_ok;

#[test]
fn free_source_format_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF1.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fixed_source_format_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FIXED\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF2.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_switch_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF3.\n>>SOURCE FORMAT FIXED\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn free_format_with_data_division_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fixed_format_with_data_division_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FIXED\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_free_with_compute_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9.\nPROCEDURE DIVISION.\n    COMPUTE X = 1 + 1.\n    STOP RUN.",
    );
}

#[test]
fn source_format_fixed_with_move_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FIXED\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9.\nPROCEDURE DIVISION.\n    MOVE 1 TO X.\n    STOP RUN.",
    );
}

#[test]
fn source_format_double_switch_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF8.\n>>SOURCE FORMAT FIXED\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9.\n>>SOURCE FORMAT FREE\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_with_comment_line_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF9.\n*> comment\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn source_format_with_copy_compiles() {
    compile_ok(
        ">>SOURCE FORMAT FREE\nIDENTIFICATION DIVISION.\nPROGRAM-ID. SF10.\nPROCEDURE DIVISION.\n    COPY COMMON-DEFS.\n    STOP RUN.",
    );
}
