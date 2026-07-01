use super::helpers::compile_ok;

#[test]
fn copy_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY1.\nPROCEDURE DIVISION.\n    COPY COMMON-DEFS.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_identifier_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY2.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_pseudotext_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY3.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC REPLACING ==OLD-NAME== BY ==NEW-NAME==.\n    STOP RUN.",
    );
}

#[test]
fn replace_off_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY4.\nPROCEDURE DIVISION.\n    REPLACE ==A== BY ==B==.\n    REPLACE OFF.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_two_names_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY5.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC REPLACING OLD-A BY NEW-A OLD-B BY NEW-B.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_literal_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY6.\nPROCEDURE DIVISION.\n    COPY TEXT-BOOK REPLACING \"OLD\" BY \"NEW\".\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_partial_words_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY7.\nPROCEDURE DIVISION.\n    COPY WORD-BOOK REPLACING ==CUST== BY ==ORD==.\n    STOP RUN.",
    );
}

#[test]
fn replace_then_copy_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY8.\nPROCEDURE DIVISION.\n    REPLACE ==A== BY ==B==.\n    COPY BASIC-BOOK.\n    REPLACE OFF.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_in_data_division_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_with_in_library_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CPY10.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC IN COMMON-LIB REPLACING OLD-NAME BY NEW-NAME.\n    STOP RUN.",
    );
}
