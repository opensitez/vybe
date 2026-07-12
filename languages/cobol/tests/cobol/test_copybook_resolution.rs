use super::helpers::compile_ok;

#[test]
fn copy_in_library_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR1.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC IN COMMON-LIB.\n    STOP RUN.",
    );
}

#[test]
fn copy_of_library_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR2.\nPROCEDURE DIVISION.\n    COPY DATE-UTILS OF COMMON-LIB.\n    STOP RUN.",
    );
}

#[test]
fn nested_copy_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR3.\nPROCEDURE DIVISION.\n    COPY OUTER-BOOK.\n    STOP RUN.",
    );
}

#[test]
fn copy_simple_book_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR4.\nPROCEDURE DIVISION.\n    COPY BASIC-BOOK.\n    STOP RUN.",
    );
}

#[test]
fn copy_replacing_simple_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR5.\nPROCEDURE DIVISION.\n    COPY BASIC-BOOK REPLACING OLD-FIELD BY NEW-FIELD.\n    STOP RUN.",
    );
}

#[test]
fn copy_in_alt_library_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR6.\nPROCEDURE DIVISION.\n    COPY ITEM-REC IN DATA-LIB.\n    STOP RUN.",
    );
}

#[test]
fn copy_of_alt_library_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR7.\nPROCEDURE DIVISION.\n    COPY ITEM-REC OF DATA-LIB.\n    STOP RUN.",
    );
}

#[test]
fn copy_pseudotext_replacing_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR8.\nPROCEDURE DIVISION.\n    COPY CUSTOMER-REC REPLACING ==CUST-ID== BY ==ORDER-ID==.\n    STOP RUN.",
    );
}

#[test]
fn copy_in_procedure_division_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR9.\nPROCEDURE DIVISION.\n    COPY PROC-BLOCK.\n    STOP RUN.",
    );
}

#[test]
fn copy_with_two_replacements_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. CBR10.\nPROCEDURE DIVISION.\n    COPY RECORD-DEF REPLACING OLD-A BY NEW-A OLD-B BY NEW-B.\n    STOP RUN.",
    );
}
