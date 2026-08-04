*> vybe-test: cobol/screen_section/screen_section_multiple_fields
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-first  PIC X(15).
       01 ws-last   PIC X(15).
       01 ws-age    PIC 99.
       SCREEN SECTION.
       01 entry-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "First Name: ".
           05 LINE 1 COLUMN 13 PIC X(15) USING ws-first.
           05 LINE 2 COLUMN 1 VALUE "Last Name:  ".
           05 LINE 2 COLUMN 13 PIC X(15) USING ws-last.
           05 LINE 3 COLUMN 1 VALUE "Age: ".
           05 LINE 3 COLUMN 6  PIC 99    USING ws-age.
       PROCEDURE DIVISION.
           MOVE "John" TO ws-first
           MOVE "Doe"  TO ws-last
           MOVE 30 TO ws-age
           DISPLAY entry-screen
           STOP RUN.

