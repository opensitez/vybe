*> vybe-test: cobol/data_group_level/data_group_three_level_nesting
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 LEVEL1.
   05 LEVEL2.
      10 LEVEL3 PIC X(4) VALUE "DEEP".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY LEVEL3.
    MOVE SPACES TO WS-VYBE-L
    STRING LEVEL3 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DEEP"
        DISPLAY "FAIL: want [DEEP] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

