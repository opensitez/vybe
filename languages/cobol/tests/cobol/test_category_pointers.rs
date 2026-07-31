use crate::helpers;

#[test]
fn test_pointers_set_address_of() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-ADDRESS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-VAL PIC X(5) VALUE "HELLO".
       01 WS-PTR USAGE POINTER.
       LINKAGE SECTION.
       01 LK-VAL PIC X(5).
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF WS-VAL.
           SET ADDRESS OF LK-VAL TO WS-PTR.
           DISPLAY LK-VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HELLO"]);
}

#[test]
fn test_pointers_set_null() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-NULL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO NULL.
           IF WS-PTR = NULL
              DISPLAY "IS NULL"
           END-IF.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["IS NULL"]);
}

#[test]
fn test_pointers_allocate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-ALLOC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-PTR USAGE POINTER.
       LINKAGE SECTION.
       01 LK-VAL PIC X(10).
       PROCEDURE DIVISION.
           ALLOCATE LK-VAL RETURNING WS-PTR.
           SET ADDRESS OF LK-VAL TO WS-PTR.
           MOVE "ALLOCATED" TO LK-VAL.
           DISPLAY LK-VAL.
           FREE WS-PTR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ALLOCATED "]);
}

#[test]
fn test_pointers_chain_address_of() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-CHAIN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-SRC PIC X(5) VALUE "HELLO".
       01 WS-DST PIC X(5).
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF WS-SRC.
           SET ADDRESS OF WS-DST TO WS-PTR.
           IF WS-DST = WS-SRC
               DISPLAY "POINTER COPY".
           END-IF.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["POINTER COPY"]);
}

#[test]
fn test_pointers_null_after_set() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-NULL-2.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO NULL.
           IF WS-PTR = NULL
              DISPLAY "NULL".
           END-IF.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["NULL"]);
}

#[test]
fn test_pointers_length_of_compiles() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-LEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-TEXT PIC X(12) VALUE "HELLO WORLD!".
       01 WS-LEN PIC 9(2).
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO LENGTH OF WS-TEXT.
           STOP RUN.
    "#;
    let _ = helpers::compile_ok(src);
}

#[test]
fn test_pointers_allocate_and_free() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-FREE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-PTR USAGE POINTER.
       01 WS-TGT PIC X(4).
       PROCEDURE DIVISION.
           ALLOCATE WS-TGT RETURNING WS-PTR.
           IF WS-PTR = NULL
               DISPLAY "NO"
           ELSE
               DISPLAY "YES"
           END-IF.
           FREE WS-PTR.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["YES"]);
}

#[test]
fn test_pointers_chain_with_table_entry() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-TBL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-BUF.
           05 ENTRIES OCCURS 2 TIMES PIC X(4) VALUE 'AAAA'.
       01 WS-PTR USAGE POINTER.
       01 WS-VIEW PIC X(4).
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF ENTRIES(2).
           SET ADDRESS OF WS-VIEW TO WS-PTR.
           DISPLAY WS-VIEW.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AAAA"]);
}
