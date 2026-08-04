*> vybe-test: cobol/screen_section/screen_section_foreground_background
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 colored-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1
              VALUE "Error!"
              FOREGROUND-COLOR 4
              BACKGROUND-COLOR 0.
       PROCEDURE DIVISION.
           DISPLAY colored-screen
           STOP RUN.

