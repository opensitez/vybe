*> vybe-test: cobol/level_66_78/renames_numeric_fields
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 date-record.
           05 date-year  PIC 9(4).
           05 date-month PIC 99.
           05 date-day   PIC 99.
       66 date-yymmdd RENAMES date-year THRU date-day.
       PROCEDURE DIVISION.
           MOVE 2024 TO date-year
           MOVE 12   TO date-month
           MOVE 25   TO date-day
           DISPLAY date-yymmdd
           STOP RUN.

