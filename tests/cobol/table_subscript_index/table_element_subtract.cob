*> vybe-test: cobol/table_subscript_index/table_element_subtract
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 200 TO E(2).
    SUBTRACT 75 FROM E(2).
    DISPLAY E(2).
    MOVE SPACES TO WS-VYBE-L
    STRING E(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "125"
        DISPLAY "FAIL: want [125] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

