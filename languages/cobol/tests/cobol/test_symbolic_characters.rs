use super::helpers::compile_ok;

#[test]
fn symbolic_characters_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM1.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS EURO IS 128.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_multiple_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM2.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS C1 IS 1 C2 IS 2 C3 IS 3.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_in_display_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM3.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS S-A IS 65.\nPROCEDURE DIVISION.\n    DISPLAY S-A.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_single_letter_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM4.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS A-SYM IS 65.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_two_names_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM5.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS S1 IS 10 S2 IS 11.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_four_names_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM6.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS S1 IS 1 S2 IS 2 S3 IS 3 S4 IS 4.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_with_display_usage_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM7.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS TAB-SYM IS 9.\nPROCEDURE DIVISION.\n    DISPLAY TAB-SYM.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_with_working_storage_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM8.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS X-SYM IS 88.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X.\nPROCEDURE DIVISION.\n    MOVE X-SYM TO CH.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_with_multiple_display_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM9.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS C1 IS 1 C2 IS 2.\nPROCEDURE DIVISION.\n    DISPLAY C1 C2.\n    STOP RUN.",
    );
}

#[test]
fn symbolic_characters_with_sectioned_program_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SYM10.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    SYMBOLIC CHARACTERS NL IS 10.\nPROCEDURE DIVISION.\nMAIN SECTION.\n    DISPLAY NL.\n    STOP RUN.",
    );
}
