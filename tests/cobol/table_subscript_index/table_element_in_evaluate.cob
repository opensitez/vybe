*> vybe-test: cobol/table_subscript_index/table_element_in_evaluate
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "B" TO E(2).
    EVALUATE E(2)
        WHEN "A" DISPLAY "ALPHA"
        WHEN "B" DISPLAY "BETA"
        WHEN OTHER DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALPHA" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALPHA"
        DISPLAY "FAIL: want [ALPHA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

