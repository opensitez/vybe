*> vybe-test: cobol/repository/repository_function_all_intrinsic_string
*> origin: languages/cobol/tests/cobol/test_repository.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           FUNCTION ALL INTRINSIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-src  PIC X(20) VALUE "  hello world  ".
       01 ws-trim PIC X(20).
       01 ws-up   PIC X(20).
       01 ws-rev  PIC X(20).
       PROCEDURE DIVISION.
           MOVE TRIM(ws-src)         TO ws-trim
           MOVE UPPER-CASE(ws-src)   TO ws-up
           MOVE REVERSE(ws-src)      TO ws-rev
           DISPLAY ws-trim
           DISPLAY ws-up
           DISPLAY ws-rev
           STOP RUN.

