use super::helpers::compile_ok;

// ═══════════════════════════════════════════════════════════
// COBOL 2023: JSON and XML processing
// Tests for JSON GENERATE/PARSE beyond test_io_and_misc.rs basics.
// ═══════════════════════════════════════════════════════════

#[test]
fn json_generate_basic() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PERSON.
   05 WS-NAME PIC X(20) VALUE "John".
   05 WS-AGE PIC 9(3) VALUE 30.
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-PERSON.
    DISPLAY WS-JSON.
    STOP RUN.
"#,
    );
}

#[test]
fn json_generate_name_override() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA.
   05 WS-FIRST-NAME PIC X(20) VALUE "Jane".
   05 WS-LAST-NAME PIC X(20) VALUE "Doe".
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-DATA
        NAME OF WS-FIRST-NAME IS "firstName"
        NAME OF WS-LAST-NAME IS "lastName".
    DISPLAY WS-JSON.
    STOP RUN.
"#,
    );
}

#[test]
fn json_generate_omit_field() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA.
   05 WS-NAME PIC X(20) VALUE "John".
   05 WS-INTERNAL PIC X(10) VALUE "secret".
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-DATA
        NAME OF WS-INTERNAL IS OMITTED.
    DISPLAY WS-JSON.
    STOP RUN.
"#,
    );
}

#[test]
fn json_parse_basic() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-JSON PIC X(100) VALUE '{"name":"Alice","age":25}'.
01 WS-DATA.
   05 WS-NAME PIC X(20).
   05 WS-AGE PIC 9(3).
PROCEDURE DIVISION.
    JSON PARSE WS-JSON INTO WS-DATA.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.
"#,
    );
}

#[test]
fn json_generate_nested() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ORDER.
   05 WS-ORDER-ID PIC 9(5) VALUE 12345.
   05 WS-CUSTOMER.
      10 WS-CUST-NAME PIC X(20) VALUE "Bob".
      10 WS-CUST-EMAIL PIC X(30) VALUE "bob@test.com".
   05 WS-TOTAL PIC 9(7)V99 VALUE 99.99.
01 WS-JSON PIC X(500).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-ORDER.
    DISPLAY WS-JSON.
    STOP RUN.
"#,
    );
}

#[test]
fn json_roundtrip() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD.
   05 WS-ID PIC 9(5) VALUE 100.
   05 WS-DESC PIC X(20) VALUE "Widget".
01 WS-JSON PIC X(200).
01 WS-RECORD2.
   05 WS-ID2 PIC 9(5).
   05 WS-DESC2 PIC X(20).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-RECORD.
    JSON PARSE WS-JSON INTO WS-RECORD2.
    DISPLAY WS-ID2.
    DISPLAY WS-DESC2.
    STOP RUN.
"#,
    );
}
