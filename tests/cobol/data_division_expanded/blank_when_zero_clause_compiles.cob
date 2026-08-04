*> vybe-test: cobol/data_division_expanded/blank_when_zero_clause_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(3) BLANK WHEN ZERO VALUE 0.
01 WS-TMP PIC X(3).
01 WS-OUT PIC X.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-NUM TO WS-TMP.
    IF WS-TMP = "   "
        MOVE 'Y' TO WS-OUT
    ELSE
        MOVE 'N' TO WS-OUT
    END-IF.
    DISPLAY WS-OUT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-OUT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

