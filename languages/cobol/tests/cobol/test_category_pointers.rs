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
