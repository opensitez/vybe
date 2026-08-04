*> vybe-test: cobol/data_group_level/data_group_group_level_condition
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 KEY.
   05 K1 PIC X(2) VALUE "AB".
   05 K2 PIC X(2) VALUE "CD".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF KEY = "ABCD"
        DISPLAY "MATCH"
    ELSE
        DISPLAY "NO MATCH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MATCH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MATCH"
        DISPLAY "FAIL: want [MATCH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

