*> vybe-test: cobol/figurative_constants/low_values_initialize_field
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-buf PIC X(20) VALUE LOW-VALUES.
       PROCEDURE DIVISION.
           DISPLAY "initialized"
           STOP RUN.

