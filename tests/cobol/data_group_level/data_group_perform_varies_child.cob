*> vybe-test: cobol/data_group_level/data_group_perform_varies_child
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DATA-REC.
   05 ITER-CNT PIC 9(2) VALUE 0.
   05 ITER-SUM PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING ITER-CNT FROM 1 BY 1 UNTIL ITER-CNT > 5
        ADD ITER-CNT TO ITER-SUM
    END-PERFORM.
    DISPLAY ITER-SUM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING ITER-SUM DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "15"
                DISPLAY "FAIL at 1 want [15] got [" WS-VYBE-L "]"
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

