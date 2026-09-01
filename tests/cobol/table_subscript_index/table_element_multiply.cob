*> vybe-test: cobol/table_subscript_index/table_element_multiply
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(4) OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 12 TO E(3).
    MULTIPLY 4 BY E(3).
    DISPLAY E(3).
    MOVE SPACES TO WS-VYBE-L
    STRING E(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0048"
        DISPLAY "FAIL: want [0048] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

