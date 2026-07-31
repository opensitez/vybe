use super::helpers::{compile_ok, run_prints};

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

#[test]
fn class_clause_runtime_vowel_check() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS11.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS VOWEL-CLASS IS \"A\" \"E\" \"I\" \"O\" \"U\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"E\".\nPROCEDURE DIVISION.\n    IF CH IS VOWEL-CLASS\n        DISPLAY \"V\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["V"]);
}

#[test]
fn class_clause_runtime_lowercase_condition() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS12.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS DIGIT-CLASS IS \"0\" THRU \"9\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-NUM PIC X VALUE \"7\".\n01 WS-CHAR PIC X VALUE \"A\".\nPROCEDURE DIVISION.\n    IF WS-NUM IS DIGIT-CLASS\n        DISPLAY \"NUM\"\n    END-IF\n    IF WS-CHAR IS NOT DIGIT-CLASS\n        DISPLAY \"NOTNUM\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["NUM", "NOTNUM"]);
}

#[test]
fn class_clause_runtime_evaluate() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS13.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS FLAG-CLASS IS \"Y\" \"N\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"Y\".\n01 RES PIC X.\nPROCEDURE DIVISION.\n    EVALUATE TRUE\n        WHEN CH IS FLAG-CLASS DISPLAY \"F\"\n        WHEN OTHER DISPLAY \"O\"\n    END-EVALUATE.\n    STOP RUN.",
    );
    assert_eq!(out, vec!["F"]);
}

#[test]
fn class_clause_runtime_multiple_classes() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS14.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS DIGITS IS \"0\" THRU \"9\".\n    CLASS LETTERS IS \"A\" THRU \"Z\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"Q\".\nPROCEDURE DIVISION.\n    IF CH IS LETTERS\n        DISPLAY \"LETTER\"\n    END-IF\n    IF CH IS NOT DIGITS\n        DISPLAY \"NOTDIG\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["LETTER", "NOTDIG"]);
}

#[test]
fn class_clause_runtime_space_class() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS15.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS SPACE-CLASS IS SPACE.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE SPACE.\nPROCEDURE DIVISION.\n    IF CH IS SPACE-CLASS\n        DISPLAY \"SPACE\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["SPACE"]);
}

#[test]
fn class_clause_zero_and_space_mix_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS16.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS ZERO-CLASS IS ZERO.\n    CLASS SPACE-CLASS IS SPACE.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn class_clause_runtime_zero_check() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS17.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS ZERO-CLASS IS ZERO.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X.\nPROCEDURE DIVISION.\n    MOVE ZERO TO CH\n    IF CH IS ZERO-CLASS\n        DISPLAY \"ZERO\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn class_clause_runtime_not_class_false_path() {
    let out = run_prints(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CLS18.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    CLASS LETTER-CLASS IS \"A\" THRU \"Z\".\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 CH PIC X VALUE \"1\".\nPROCEDURE DIVISION.\n    IF CH IS NOT LETTER-CLASS\n        DISPLAY \"NOT_LETTER\"\n    END-IF\n    STOP RUN.",
    );
    assert_eq!(out, vec!["NOT_LETTER"]);
}
