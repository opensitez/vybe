*> vybe-test: cobol/data_division_expanded/alphanumeric_editted_item_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(12) VALUE "ABC".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "123" TO WS-TXT.
    DISPLAY WS-TXT.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-TXT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "123"
        DISPLAY "FAIL: want [123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

