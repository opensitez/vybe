*> vybe-test: cobol/linage/linage_basic
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT print-file ASSIGN TO "report.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD print-file
           LINAGE IS 20 LINES.
       01 print-rec PIC X(80).
       PROCEDURE DIVISION.
           OPEN OUTPUT print-file
           MOVE "Hello, Report!" TO print-rec
           WRITE print-rec
           CLOSE print-file
           STOP RUN.

