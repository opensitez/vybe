*> vybe-test: cobol/screen_section/screen_section_blank
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 main-screen.
           05 BLANK SCREEN.
       PROCEDURE DIVISION.
           DISPLAY main-screen
           STOP RUN.

