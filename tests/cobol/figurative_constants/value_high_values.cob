*> vybe-test: cobol/figurative_constants/value_high_values
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-max-key PIC X(10) VALUE HIGH-VALUES.
       01 ws-min-key PIC X(10) VALUE LOW-VALUES.
       01 ws-blank   PIC X(10) VALUE SPACES.
       01 ws-zero    PIC 9(5)  VALUE ZEROS.
       PROCEDURE DIVISION.
           DISPLAY "initialized"
           STOP RUN.

