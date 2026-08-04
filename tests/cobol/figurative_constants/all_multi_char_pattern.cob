*> vybe-test: cobol/figurative_constants/all_multi_char_pattern
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-banner PIC X(12).
       PROCEDURE DIVISION.
           MOVE ALL "AB" TO ws-banner
           DISPLAY ws-banner
           STOP RUN.

