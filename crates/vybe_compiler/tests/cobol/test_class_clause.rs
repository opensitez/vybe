use super::helpers::compile_ok;

#[test]
fn special_names_class_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS1.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS DIGIT-CLASS IS \"0\" THRU \"9\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn if_class_condition_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS2.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS VOWEL-CLASS IS \"A\" \"E\" \"I\" \"O\" \"U\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"A\".\nPROCEDURE DIVISION.\n    IF CH IS VOWEL-CLASS DISPLAY \"V\" END-IF.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_multiple_literals_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS3.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS HEX-CLASS IS \"A\" THRU \"F\" \"0\" THRU \"9\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_lowercase_range_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS4.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS LOWER-CLASS IS \"a\" THRU \"z\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_uppercase_range_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS5.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS UPPER-CLASS IS \"A\" THRU \"Z\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_alphanumeric_set_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS6.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS ID-CLASS IS \"A\" THRU \"Z\" \"0\" THRU \"9\" \"-\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_if_negative_example_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS7.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS STAR-CLASS IS \"*\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"*\".\nPROCEDURE DIVISION.\n    IF CH IS STAR-CLASS DISPLAY \"Y\" END-IF.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_multiple_named_classes_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS8.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS DIGITS IS \"0\" THRU \"9\".\n    CLASS LETTERS IS \"A\" THRU \"Z\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_whitespace_class_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS9.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS SPACE-CLASS IS SPACE.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_condition_in_evaluate_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS10.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS FLAG-CLASS IS \"Y\" \"N\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"Y\".\nPROCEDURE DIVISION.\n    EVALUATE TRUE WHEN CH IS FLAG-CLASS DISPLAY \"F\" WHEN OTHER DISPLAY \"O\" END-EVALUATE.\n    STOP RUN.",
    );
}
