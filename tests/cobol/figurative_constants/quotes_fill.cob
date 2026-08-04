*> vybe-test: cobol/figurative_constants/quotes_fill
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(5).
       PROCEDURE DIVISION.
           MOVE QUOTES TO ws-field
           DISPLAY ws-field
           STOP RUN.

