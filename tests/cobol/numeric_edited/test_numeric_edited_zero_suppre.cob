*> vybe-test: cobol/numeric_edited/test_numeric_edited_zero_suppress
*> origin: languages/cobol/tests/cobol/test_numeric_edited.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC 9(3) VALUE 42.
01 WS-EDIT PIC Z(3) VALUE ZERO.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE WS-VAL TO WS-EDIT.
    DISPLAY WS-EDIT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-EDIT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = " 42"
        DISPLAY "FAIL: want [ 42] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

