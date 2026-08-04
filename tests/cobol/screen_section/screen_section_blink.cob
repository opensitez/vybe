*> vybe-test: cobol/screen_section/screen_section_blink
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 alert-screen.
           05 LINE 12 COLUMN 30 VALUE "ALERT!" BLINK.
       PROCEDURE DIVISION.
           DISPLAY alert-screen
           STOP RUN.

