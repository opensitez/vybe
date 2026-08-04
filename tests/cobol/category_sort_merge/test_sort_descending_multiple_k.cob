*> vybe-test: cobol/category_sort_merge/test_sort_descending_multiple_keys
*> origin: languages/cobol/tests/cobol/test_category_sort_merge.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-DESC.
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
          05 KEY-1 PIC 9(2).
          05 KEY-2 PIC 9(2).
       FD IN-FILE.
       01 IN-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           SORT SORT-WORK ON ASCENDING KEY KEY-1
                          ON DESCENDING KEY KEY-2
              USING IN-FILE
              GIVING OUT-FILE.
           DISPLAY "SORT DESC PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "SORT DESC PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SORT DESC PARSED"
        DISPLAY "FAIL: want [SORT DESC PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

