*> vybe-test: cobol/figurative_constants/all_star_fill
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-border PIC X(40).
       PROCEDURE DIVISION.
           MOVE ALL "*" TO ws-border
           DISPLAY ws-border
           STOP RUN.

