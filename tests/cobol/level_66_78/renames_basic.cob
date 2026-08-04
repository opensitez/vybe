*> vybe-test: cobol/level_66_78/renames_basic
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 address-record.
           05 street    PIC X(30).
           05 city      PIC X(20).
           05 state     PIC XX.
           05 zip-code  PIC X(10).
       66 city-state RENAMES city THRU state.
       PROCEDURE DIVISION.
           MOVE "Springfield" TO city
           MOVE "IL"          TO state
           DISPLAY city-state
           STOP RUN.

