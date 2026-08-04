*> vybe-test: cobol/screen_section/screen_section_auto_tab
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-code PIC X(5).
       SCREEN SECTION.
       01 code-screen.
           05 LINE 5 COLUMN 10 VALUE "Code: ".
           05 LINE 5 COLUMN 16 PIC X(5) USING ws-code AUTO.
       PROCEDURE DIVISION.
           DISPLAY code-screen
           STOP RUN.

