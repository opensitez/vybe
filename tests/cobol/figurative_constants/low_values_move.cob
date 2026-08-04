*> vybe-test: cobol/figurative_constants/low_values_move
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-init PIC X(10).
       PROCEDURE DIVISION.
           MOVE LOW-VALUES TO ws-init
           DISPLAY "low values set"
           STOP RUN.

