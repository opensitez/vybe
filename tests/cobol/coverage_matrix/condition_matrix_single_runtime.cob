*> vybe-test: cobol/coverage_matrix/condition_matrix_single_runtime
*> origin: languages/cobol/tests/cobol/test_coverage_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SAMPLE2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 A PIC 9(2) VALUE 7.
01 B PIC 9(2) VALUE 7.
01 FLAG PIC X(1) VALUE "N".
PROCEDURE DIVISION.
    IF A = B MOVE "Y" TO FLAG END-IF.
    DISPLAY FLAG
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING FLAG DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

