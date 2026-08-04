*> vybe-test: cobol/data_group_level/data_group_filler_invisible_in_display
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 FILLER PIC X(3) VALUE "XXX".
   05 DATA-PART PIC X(5) VALUE "HELLO".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY DATA-PART.
    MOVE SPACES TO WS-VYBE-L
    STRING DATA-PART DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO"
        DISPLAY "FAIL: want [HELLO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

