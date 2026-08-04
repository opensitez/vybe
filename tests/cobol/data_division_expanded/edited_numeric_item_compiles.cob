*> vybe-test: cobol/data_division_expanded/edited_numeric_item_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ED PIC ZZ9.99 VALUE 12.34.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 7.5 TO WS-ED.
    DISPLAY WS-ED.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ED DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7.5"
        DISPLAY "FAIL: want [7.5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

