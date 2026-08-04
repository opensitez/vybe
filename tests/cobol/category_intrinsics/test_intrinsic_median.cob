*> vybe-test: cobol/category_intrinsics/test_intrinsic_median
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. MEDIAN-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 TBL-DATA.
          05 VALS OCCURS 5 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 50 TO VALS(1).
           MOVE 10 TO VALS(2).
           MOVE 30 TO VALS(3).
           MOVE 20 TO VALS(4).
           MOVE 40 TO VALS(5).
           COMPUTE RES = FUNCTION MEDIAN(ALL VALS).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "30"
        DISPLAY "FAIL: want [30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

