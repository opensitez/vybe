*> vybe-test: cobol/table_subscript_index/table_subscript_integer_literal
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(2) OCCURS 5 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 77 TO E(3).
    DISPLAY E(3).
    MOVE SPACES TO WS-VYBE-L
    STRING E(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "77"
        DISPLAY "FAIL: want [77] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

