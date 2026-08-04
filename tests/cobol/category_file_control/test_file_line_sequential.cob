*> vybe-test: cobol/category_file_control/test_file_line_sequential
*> origin: languages/cobol/tests/cobol/test_category_file_control.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-LINE.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT LSEQ-FILE ASSIGN TO "lines.txt"
           ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD LSEQ-FILE.
       01 LREC PIC X(50).
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           DISPLAY "LINE SEQ PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "LINE SEQ PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LINE SEQ PARSED"
        DISPLAY "FAIL: want [LINE SEQ PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

