*> vybe-test: cobol/linage/linage_with_variable
*> origin: languages/cobol/tests/cobol/test_linage.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT var-rpt ASSIGN TO "var.txt"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD var-rpt
           LINAGE IS ws-page-lines LINES
           WITH FOOTING AT ws-footing-line
           LINES AT TOP ws-top-margin
           LINES AT BOTTOM ws-bot-margin.
       01 var-line PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-page-lines    PIC 99 VALUE 55.
       01 ws-footing-line  PIC 99 VALUE 53.
       01 ws-top-margin    PIC 9  VALUE 2.
       01 ws-bot-margin    PIC 9  VALUE 3.
       PROCEDURE DIVISION.
           OPEN OUTPUT var-rpt
           MOVE "Variable linage test" TO var-line
           WRITE var-line
           CLOSE var-rpt
           STOP RUN.

