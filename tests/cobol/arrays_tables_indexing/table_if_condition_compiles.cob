*> vybe-test: cobol/arrays_tables_indexing/table_if_condition_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9(2) OCCURS 3 TIMES.
01 I PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 2 TO I.
    MOVE 5 TO T(2).
    IF T(I) = 5 DISPLAY "MATCH" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MATCH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MATCH"
        DISPLAY "FAIL: want [MATCH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

