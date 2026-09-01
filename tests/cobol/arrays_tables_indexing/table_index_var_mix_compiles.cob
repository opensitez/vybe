*> vybe-test: cobol/arrays_tables_indexing/table_index_var_mix_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9(2) OCCURS 3 TIMES.
01 I PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 9 TO T(I).
    DISPLAY T(I).
    MOVE SPACES TO WS-VYBE-L
    STRING T(I) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "09"
        DISPLAY "FAIL: want [09] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

