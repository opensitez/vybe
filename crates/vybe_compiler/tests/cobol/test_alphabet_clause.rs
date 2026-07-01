use super::helpers::compile_ok;

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
fn alphabet_clause_named_for_collating_sequence_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ALPHA10.\nENVIRONMENT DIVISION.\nCONFIGURATION SECTION.\nSPECIAL-NAMES.\n    ALPHABET MY-ALPHA IS STANDARD-1.\n    COLLATING SEQUENCE IS MY-ALPHA.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}
