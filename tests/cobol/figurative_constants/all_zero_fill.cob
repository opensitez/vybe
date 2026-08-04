*> vybe-test: cobol/figurative_constants/all_zero_fill
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(10).
       PROCEDURE DIVISION.
           MOVE ALL "0" TO ws-field
           DISPLAY ws-field
           STOP RUN.

