*> vybe-test: cobol/sort_advanced/release_from_working_storage
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sf ASSIGN TO "sf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sf.
       01 srec.
           05 sk   PIC X(5).
           05 sval PIC 9(5).
       WORKING-STORAGE SECTION.
       01 ws-key PIC X(5) VALUE "AKEY".
       01 ws-val PIC 9(5) VALUE 42.
       PROCEDURE DIVISION.
           SORT sf ON ASCENDING KEY sk
               INPUT PROCEDURE IS fill-sort
               GIVING "out.dat"
           STOP RUN.
       fill-sort SECTION.
           MOVE ws-key TO sk
           MOVE ws-val TO sval
           RELEASE srec
           MOVE "BKEY" TO sk
           MOVE 99 TO sval
           RELEASE srec FROM srec.

