*> vybe-test: cobol/screen_section/screen_section_required
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-id PIC 9(8).
       SCREEN SECTION.
       01 id-screen.
           05 LINE 3 COLUMN 5 VALUE "ID: ".
           05 LINE 3 COLUMN 9 PIC 9(8) USING ws-id REQUIRED.
       PROCEDURE DIVISION.
           MOVE 12345678 TO ws-id
           DISPLAY id-screen
           STOP RUN.

