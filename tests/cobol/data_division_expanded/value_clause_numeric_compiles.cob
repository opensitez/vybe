*> vybe-test: cobol/data_division_expanded/value_clause_numeric_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(3) VALUE 100.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY WS-NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "100"
        DISPLAY "FAIL: want [100] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

