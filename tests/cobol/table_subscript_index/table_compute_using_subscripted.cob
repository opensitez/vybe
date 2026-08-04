*> vybe-test: cobol/table_subscript_index/table_compute_using_subscripted_field
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 N PIC 9(3) OCCURS 3 TIMES.
01 R PIC 9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 5 TO N(1). MOVE 10 TO N(2). MOVE 15 TO N(3).
    COMPUTE R = N(1) * N(2) + N(3).
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "65"
        DISPLAY "FAIL: want [65] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

