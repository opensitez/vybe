*> vybe-test: cobol/data_group_level/data_group_redisplay_after_move_into_child
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PAIR.
   05 LEFT PIC X(3) VALUE "AAA".
   05 RIGHT PIC X(3) VALUE "BBB".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "ZZZ" TO LEFT.
    DISPLAY PAIR.
    MOVE SPACES TO WS-VYBE-L
    STRING PAIR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZZZBBB"
        DISPLAY "FAIL: want [ZZZBBB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

