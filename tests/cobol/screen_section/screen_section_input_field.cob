*> vybe-test: cobol/screen_section/screen_section_input_field
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-name PIC X(20).
       SCREEN SECTION.
       01 name-screen.
           05 BLANK SCREEN.
           05 LINE 2 COLUMN 5 VALUE "Name: ".
           05 LINE 2 COLUMN 12 PIC X(20) USING ws-name.
       PROCEDURE DIVISION.
           MOVE "Alice" TO ws-name
           DISPLAY name-screen
           STOP RUN.

