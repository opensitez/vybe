*> vybe-test: cobol/figurative_constants/all_numeric_pattern
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-mask PIC X(8).
       PROCEDURE DIVISION.
           MOVE ALL "12" TO ws-mask
           DISPLAY ws-mask
           STOP RUN.

