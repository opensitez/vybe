*> vybe-test: cobol/category_sort_merge/test_sort_input_output_procedure
*> origin: languages/cobol/tests/cobol/test_category_sort_merge.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-PROC.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SORT-WORK ASSIGN TO "work.dat".
       DATA DIVISION.
       FILE SECTION.
       SD SORT-WORK.
       01 WORK-REC.
          05 SORT-KEY PIC 9(4).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
       MAIN SECTION.
           SORT SORT-WORK ON ASCENDING KEY SORT-KEY
              INPUT PROCEDURE IS IN-PROC
              OUTPUT PROCEDURE IS OUT-PROC.
           DISPLAY "SORT PROCS PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "SORT PROCS PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SORT PROCS PARSED"
        DISPLAY "FAIL: want [SORT PROCS PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.
       IN-PROC SECTION.
           EXIT.
       OUT-PROC SECTION.
           EXIT.

