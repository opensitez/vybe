*> vybe-test: cobol/data_group_level/data_group_four_levels_deep
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 L1.
   05 L2.
      10 L3.
         15 L4 PIC 9(4) VALUE 4321.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY L4.
    MOVE SPACES TO WS-VYBE-L
    STRING L4 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "4321"
        DISPLAY "FAIL: want [4321] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

