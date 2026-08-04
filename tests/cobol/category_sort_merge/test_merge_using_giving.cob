*> vybe-test: cobol/category_sort_merge/test_merge_using_giving
*> origin: languages/cobol/tests/cobol/test_category_sort_merge.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. MERGE-UG.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT MERGE-WORK ASSIGN TO "work.dat".
           SELECT IN-1 ASSIGN TO "in1.dat".
           SELECT IN-2 ASSIGN TO "in2.dat".
           SELECT OUT-FILE ASSIGN TO "out.dat".
       DATA DIVISION.
       FILE SECTION.
       SD MERGE-WORK.
       01 WORK-REC.
          05 MERGE-KEY PIC 9(4).
       FD IN-1.
       01 IN1-REC PIC X(10).
       FD IN-2.
       01 IN2-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           MERGE MERGE-WORK ON ASCENDING KEY MERGE-KEY
              USING IN-1 IN-2
              GIVING OUT-FILE.
           DISPLAY "MERGE PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "MERGE PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MERGE PARSED"
        DISPLAY "FAIL: want [MERGE PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

