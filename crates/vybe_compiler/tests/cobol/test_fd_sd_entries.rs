use super::helpers::compile_ok;

#[test]
fn fd_entry_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT F ASSIGN TO \"f.dat\".\nDATA DIVISION.\nFILE SECTION.\nFD F RECORD CONTAINS 80 CHARACTERS.\n01 R PIC X(80).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn sd_entry_basic_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nSD SORT-FILE.\n01 SORT-REC.\n   05 SORT-KEY PIC 9(5).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_block_contains_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F BLOCK CONTAINS 5 RECORDS.\n01 R PIC X(80).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_label_records_standard_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F LABEL RECORDS ARE STANDARD.\n01 R PIC X(20).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_label_records_omitted_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F LABEL RECORDS ARE OMITTED.\n01 R PIC X(20).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_data_records_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F DATA RECORDS ARE R1 R2.\n01 R1 PIC X(20).\n01 R2 PIC X(30).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_multiple_01_entries_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F.\n01 R1 PIC X(20).\n01 R2 PIC X(30).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn fd_external_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nFD F EXTERNAL.\n01 R PIC X(20).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn sd_record_contains_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nSD SORT-FILE RECORD CONTAINS 50 CHARACTERS.\n01 SORT-REC PIC X(50).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn sd_data_records_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nFILE SECTION.\nSD SORT-FILE DATA RECORDS ARE SORT-REC.\n01 SORT-REC PIC X(40).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}