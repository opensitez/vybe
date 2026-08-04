*> vybe-test: cobol/data_group_level/data_group_display_as_whole_group
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PAIR.
   05 PART-A PIC X(3) VALUE "ABC".
   05 PART-B PIC X(3) VALUE "DEF".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY PAIR.
    MOVE SPACES TO WS-VYBE-L
    STRING PAIR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDEF"
        DISPLAY "FAIL: want [ABCDEF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

