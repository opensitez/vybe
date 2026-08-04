*> vybe-test: cobol/screen_section/screen_section_with_accept
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-choice PIC X VALUE SPACE.
       SCREEN SECTION.
       01 menu-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "1. Option A".
           05 LINE 2 COLUMN 1 VALUE "2. Option B".
           05 LINE 4 COLUMN 1 VALUE "Choice: ".
           05 choice-fld LINE 4 COLUMN 9 PIC X USING ws-choice.
       PROCEDURE DIVISION.
           DISPLAY menu-screen
           STOP RUN.

