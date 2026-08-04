*> vybe-test: cobol/figurative_constants/all_single_char_fill
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-line PIC X(20).
       PROCEDURE DIVISION.
           MOVE ALL "-" TO ws-line
           DISPLAY ws-line
           STOP RUN.

