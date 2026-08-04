*> vybe-test: cobol/scope/nested_if_scope_compiles
*> origin: languages/cobol/tests/cobol/test_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(2) VALUE 5.
PROCEDURE DIVISION.
    IF WS-X > 0
        IF WS-X < 10
            DISPLAY "in-range"
        END-IF
    END-IF.
    STOP RUN.

