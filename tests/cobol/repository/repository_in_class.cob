*> vybe-test: cobol/repository/repository_in_class
*> origin: languages/cobol/tests/cobol/test_repository.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

       CLASS-ID. Calculator.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-precision PIC 99 VALUE 8.
       METHOD-ID. square-root.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-result PIC 9(5)V9(8).
       LINKAGE SECTION.
       01 lk-input  PIC 9(5)V9(8).
       01 lk-result PIC 9(5)V9(8).
       PROCEDURE DIVISION USING lk-input RETURNING lk-result.
           COMPUTE lk-result = SQRT(lk-input)
           GOBACK.
       END METHOD square-root.
       END CLASS Calculator.
    STOP RUN.

