*> vybe-test: cobol/category_file_control/test_file_control_indexed
*> origin: languages/cobol/tests/cobol/test_category_file_control.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-IDX.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT IDX-FILE ASSIGN TO "idx.dat"
           ORGANIZATION IS INDEXED
           ACCESS MODE IS RANDOM
           RECORD KEY IS KEY-FLD
           ALTERNATE RECORD KEY IS ALT-KEY WITH DUPLICATES.
       DATA DIVISION.
       FILE SECTION.
       FD IDX-FILE.
       01 IDX-REC.
          05 KEY-FLD PIC X(5).
          05 ALT-KEY PIC X(5).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           DISPLAY "IDX FILE PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "IDX FILE PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "IDX FILE PARSED"
        DISPLAY "FAIL: want [IDX FILE PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

