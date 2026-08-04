*> vybe-test: cobol/linage/linage_with_top_bottom
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT page-file ASSIGN TO "page.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD page-file
           LINAGE IS 40 LINES
           WITH FOOTING AT 38
           LINES AT TOP 3
           LINES AT BOTTOM 3.
       01 page-rec PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT page-file
           MOVE "First line" TO page-rec
           WRITE page-rec
           CLOSE page-file
           STOP RUN.

