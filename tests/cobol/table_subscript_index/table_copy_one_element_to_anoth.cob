*> vybe-test: cobol/table_subscript_index/table_copy_one_element_to_another
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 5 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 123 TO E(1).
    MOVE E(1) TO E(5).
    DISPLAY E(5).
    MOVE SPACES TO WS-VYBE-L
    STRING E(5) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "123"
        DISPLAY "FAIL: want [123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

