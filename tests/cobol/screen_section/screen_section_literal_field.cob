*> vybe-test: cobol/screen_section/screen_section_literal_field
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 header-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "Welcome to COBOL".
       PROCEDURE DIVISION.
           DISPLAY header-screen
           STOP RUN.

