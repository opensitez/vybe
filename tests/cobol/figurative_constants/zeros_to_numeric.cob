*> vybe-test: cobol/figurative_constants/zeros_to_numeric
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-n PIC 9(5).
       PROCEDURE DIVISION.
           MOVE ZEROS TO ws-n
           DISPLAY ws-n
           STOP RUN.

