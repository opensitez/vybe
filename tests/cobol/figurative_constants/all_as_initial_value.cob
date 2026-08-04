*> vybe-test: cobol/figurative_constants/all_as_initial_value
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-underline PIC X(30) VALUE ALL "-".
       01 ws-dots      PIC X(20) VALUE ALL "...".
       PROCEDURE DIVISION.
           DISPLAY ws-underline
           DISPLAY ws-dots
           STOP RUN.

