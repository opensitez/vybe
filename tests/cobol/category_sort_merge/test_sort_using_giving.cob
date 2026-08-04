*> vybe-test: cobol/category_sort_merge/test_sort_using_giving
*> origin: languages/cobol/tests/cobol/test_category_sort_merge.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-UG.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SORT-WORK ASSIGN TO "work.dat".
           SELECT IN-FILE ASSIGN TO "in.dat".
           SELECT OUT-FILE ASSIGN TO "out.dat".
       DATA DIVISION.
       FILE SECTION.
       SD SORT-WORK.
       01 WORK-REC.
          05 SORT-KEY PIC 9(4).
       FD IN-FILE.
       01 IN-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           SORT SORT-WORK ON ASCENDING KEY SORT-KEY
              USING IN-FILE
              GIVING OUT-FILE.
           DISPLAY "SORT USING GIVING PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "SORT USING GIVING PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SORT USING GIVING PARSED"
        DISPLAY "FAIL: want [SORT USING GIVING PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

