*> vybe-test: cobol/category_intrinsics/test_intrinsic_range
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. RANGE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 TBL-DATA.
          05 VALS OCCURS 5 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 15 TO VALS(1).
           MOVE 05 TO VALS(2).
           MOVE 45 TO VALS(3).
           MOVE 25 TO VALS(4).
           MOVE 35 TO VALS(5).
           COMPUTE RES = FUNCTION RANGE(ALL VALS).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "40"
        DISPLAY "FAIL: want [40] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

