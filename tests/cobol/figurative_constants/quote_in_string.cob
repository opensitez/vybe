*> vybe-test: cobol/figurative_constants/quote_in_string
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-msg PIC X(30).
       PROCEDURE DIVISION.
           MOVE QUOTE TO ws-msg
           DISPLAY ws-msg
           STOP RUN.

