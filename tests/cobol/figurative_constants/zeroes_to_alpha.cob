*> vybe-test: cobol/figurative_constants/zeroes_to_alpha
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(6).
       PROCEDURE DIVISION.
           MOVE ZEROES TO ws-field
           DISPLAY ws-field
           STOP RUN.

