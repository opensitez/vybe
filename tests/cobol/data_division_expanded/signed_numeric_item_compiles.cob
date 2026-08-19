*> vybe-test: cobol/data_division_expanded/signed_numeric_item_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC S9(3) VALUE -3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE -4 TO WS-NUM.
    DISPLAY WS-NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00t"
        DISPLAY "FAIL: want [00t] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

