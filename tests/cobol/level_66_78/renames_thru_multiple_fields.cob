*> vybe-test: cobol/level_66_78/renames_thru_multiple_fields
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 employee-record.
           05 emp-first  PIC X(15).
           05 emp-middle PIC X(1).
           05 emp-last   PIC X(20).
           05 emp-dept   PIC X(10).
       66 emp-name RENAMES emp-first THRU emp-last.
       PROCEDURE DIVISION.
           MOVE "John"    TO emp-first
           MOVE "A"       TO emp-middle
           MOVE "Smith"   TO emp-last
           DISPLAY emp-name
           STOP RUN.

