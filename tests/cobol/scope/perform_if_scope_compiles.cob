*> vybe-test: cobol/scope/perform_if_scope_compiles
*> origin: languages/cobol/tests/cobol/test_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 2
        ADD 1 TO I
        IF F = 1 DISPLAY I END-IF
    END-PERFORM.
    STOP RUN.

