*> vybe-test: cobol/repository/repository_function_all_intrinsic_math
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pi      PIC 9V9(8).
       01 ws-e       PIC 9V9(8).
       01 ws-abs-val PIC 99V99.
       PROCEDURE DIVISION.
           COMPUTE ws-pi      = ACOS(-1)
           COMPUTE ws-e       = EXP(1)
           COMPUTE ws-abs-val = ABS(-3.14)
           DISPLAY ws-pi
           DISPLAY ws-e
           DISPLAY ws-abs-val
           STOP RUN.

