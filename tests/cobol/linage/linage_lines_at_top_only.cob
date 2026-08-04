*> vybe-test: cobol/linage/linage_lines_at_top_only
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT out-file ASSIGN TO "out.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD out-file
           LINAGE IS 50 LINES
           LINES AT TOP 5.
       01 out-line PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT out-file
           MOVE "Body line" TO out-line
           WRITE out-line
           CLOSE out-file
           STOP RUN.

