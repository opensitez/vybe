*> vybe-test: cobol/repository/repository_specific_function
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION SQRT
           FUNCTION ABS
           FUNCTION MOD.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 9(5)V99.
       PROCEDURE DIVISION.
           COMPUTE ws-val = SQRT(25)
           DISPLAY ws-val
           COMPUTE ws-val = ABS(-7)
           DISPLAY ws-val
           STOP RUN.

