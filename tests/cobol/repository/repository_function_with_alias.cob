*> vybe-test: cobol/repository/repository_function_with_alias
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION UPPER-CASE AS "UPPER-CASE"
           FUNCTION LOWER-CASE AS "LOWER-CASE"
           FUNCTION TRIM       AS "TRIM".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-text PIC X(20) VALUE "Hello World".
       01 ws-up   PIC X(20).
       01 ws-lo   PIC X(20).
       PROCEDURE DIVISION.
           MOVE UPPER-CASE(ws-text) TO ws-up
           MOVE LOWER-CASE(ws-text) TO ws-lo
           DISPLAY ws-up
           DISPLAY ws-lo
           STOP RUN.

