use super::helpers::{compile_ok, run_prints};

#[test]
fn alphabet_clause_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA1.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET ALPHA-1 IS STANDARD-1.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_with_literal_range_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA2.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET ALPHA-2 IS \"A\" THRU \"Z\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_national_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA3.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET ALPHA-3 IS NATIVE.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_standard2_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA4.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET ALPHA-4 IS STANDARD-2.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_lowercase_range_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA5.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET LOWER-SET IS \"a\" THRU \"z\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_digits_range_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA6.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET DIGIT-SET IS \"0\" THRU \"9\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_multiple_ranges_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA7.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET HEX-SET IS \"0\" THRU \"9\" \"A\" THRU \"F\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_with_sort_collating_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA8.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET SORT-ALPHA IS STANDARD-1.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_graphic_symbols_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA9.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET PUNCT-SET IS \".\" \",\" \";\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_compares_ordered_with_collating_sequence() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA11.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MY-ALPHA IS STANDARD-1.
    COLLATING SEQUENCE IS MY-ALPHA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X VALUE "A".
01 B PIC X VALUE "B".
PROCEDURE DIVISION.
    IF A < B
        DISPLAY "ORDERED"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["ORDERED"]);
}

#[test]
fn alphabet_clause_named_for_collating_sequence_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA10.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET MY-ALPHA IS STANDARD-1.\n    COLLATING SEQUENCE IS MY-ALPHA.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn alphabet_clause_literal_range_orders_letters() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA12.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-LOW IS "A" THRU "Z".
    COLLATING SEQUENCE IS ALPHA-LOW.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X VALUE "C".
01 B PIC X VALUE "F".
PROCEDURE DIVISION.
    IF A < B
        DISPLAY "LESS"
    ELSE
        DISPLAY "NOT-LESS"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["LESS"]);
}

#[test]
fn alphabet_clause_digits_compare() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA13.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET DIGIT-SET IS "0" THRU "9".
    COLLATING SEQUENCE IS DIGIT-SET.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X VALUE "3".
01 B PIC X VALUE "8".
PROCEDURE DIVISION.
    IF A < B
        DISPLAY "LOW"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["LOW"]);
}

#[test]
fn alphabet_clause_ascii_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA14.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-ASCII IS ASCII.
PROCEDURE DIVISION.
    DISPLAY "ASCII".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["ASCII"]);
}

#[test]
fn alphabet_clause_ebcdic_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA15.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-EBCDIC IS EBCDIC.
PROCEDURE DIVISION.
    DISPLAY "EBCDIC".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["EBCDIC"]);
}

#[test]
fn alphabet_clause_native_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA16.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-NATIVE IS NATIVE.
PROCEDURE DIVISION.
    DISPLAY "NATIVE".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["NATIVE"]);
}

#[test]
fn alphabet_clause_standard_2_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA17.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-ST2 IS STANDARD-2.
PROCEDURE DIVISION.
    DISPLAY "STD2".
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["STD2"]);
}

#[test]
fn alphabet_clause_multiple_alphabets_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA18.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET ALPHA-LOW IS "a" THRU "z".
    ALPHABET ALPHA-UPPER IS "A" THRU "Z".
    COLLATING SEQUENCE IS ALPHA-LOW.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LOW  PIC X VALUE "b".
01 WS-UPER PIC X VALUE "A".
PROCEDURE DIVISION.
    IF WS-LOW < WS-UPER
        DISPLAY "LOWER-FIRST"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["LOWER-FIRST"]);
}

#[test]
fn alphabet_clause_custom_order_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA19.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET PRIORITY-ALPHA IS "C" "B" "A" "D" THRU "F".
    COLLATING SEQUENCE IS PRIORITY-ALPHA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FIRST  PIC X VALUE "C".
01 SECOND PIC X VALUE "A".
PROCEDURE DIVISION.
    IF FIRST > SECOND
        DISPLAY "C-A"
    ELSE
        DISPLAY "NOT-C-A"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["C-A"]);
}

#[test]
fn alphabet_clause_alphanumeric_mix_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA20.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MIXED-ALPHA IS "0" THRU "9" "A" THRU "Z".
    COLLATING SEQUENCE IS MIXED-ALPHA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DIGIT  PIC X VALUE "8".
01 LETTER PIC X VALUE "A".
PROCEDURE DIVISION.
    IF DIGIT < LETTER
        DISPLAY "MIXED-ORDER"
    END-IF.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["MIXED-ORDER"]);
}
