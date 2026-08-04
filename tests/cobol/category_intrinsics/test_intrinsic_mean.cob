*> vybe-test: cobol/category_intrinsics/test_intrinsic_mean
*> origin: languages/cobol/tests/cobol/test_category_intrinsics.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. MEAN-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 TBL-DATA.
          05 VALS OCCURS 4 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 10 TO VALS(1).
           MOVE 20 TO VALS(2).
           MOVE 30 TO VALS(3).
           MOVE 40 TO VALS(4).
           COMPUTE RES = FUNCTION MEAN(ALL VALS).
           DISPLAY RES.
    MOVE SPACES TO WS-VYBE-L
    STRING RES DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "25"
        DISPLAY "FAIL: want [25] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

