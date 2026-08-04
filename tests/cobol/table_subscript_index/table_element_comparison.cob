*> vybe-test: cobol/table_subscript_index/table_element_comparison
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 100 TO E(1).
    MOVE 200 TO E(2).
    IF E(2) > E(1)
        DISPLAY "SECOND BIGGER"
    ELSE
        DISPLAY "FIRST BIGGER OR EQUAL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "SECOND BIGGER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SECOND BIGGER"
        DISPLAY "FAIL: want [SECOND BIGGER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

