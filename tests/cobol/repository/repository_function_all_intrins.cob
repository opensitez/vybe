*> vybe-test: cobol/repository/repository_function_all_intrinsic
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5)V99.
       01 ws-text   PIC X(20) VALUE "hello world".
       PROCEDURE DIVISION.
           COMPUTE ws-result = SQRT(16)
           DISPLAY ws-result
           DISPLAY UPPER-CASE(ws-text)
           STOP RUN.

