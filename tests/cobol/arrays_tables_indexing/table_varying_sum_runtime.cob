*> vybe-test: cobol/arrays_tables_indexing/table_varying_sum_runtime
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9(2) OCCURS 4 TIMES.
01 I PIC 9 VALUE 1.
01 TOT PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 1 TO T(1).
    MOVE 2 TO T(2).
    MOVE 3 TO T(3).
    MOVE 4 TO T(4).
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 4
        ADD T(I) TO TOT
    END-PERFORM.
    DISPLAY TOT.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING TOT DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "010"
                DISPLAY "FAIL at 1 want [010] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

