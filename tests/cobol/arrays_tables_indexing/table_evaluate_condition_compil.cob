*> vybe-test: cobol/arrays_tables_indexing/table_evaluate_condition_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9 OCCURS 3 TIMES.
01 I PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 2 TO I.
    MOVE 2 TO T(2).
    EVALUATE T(I)
        WHEN 1 DISPLAY "ONE"
        WHEN 2 DISPLAY "TWO"
        WHEN OTHER DISPLAY "X"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

