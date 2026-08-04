*> vybe-test: cobol/screen_section/screen_section_nested_group
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-city  PIC X(20).
       01 ws-state PIC XX.
       SCREEN SECTION.
       01 addr-screen.
           05 top-line.
               10 BLANK SCREEN.
               10 LINE 1 COLUMN 1 VALUE "Address Entry".
           05 city-grp.
               10 LINE 3 COLUMN 1 VALUE "City:  ".
               10 LINE 3 COLUMN 8 PIC X(20) USING ws-city.
               10 LINE 3 COLUMN 29 VALUE "State: ".
               10 LINE 3 COLUMN 36 PIC XX USING ws-state.
       PROCEDURE DIVISION.
           MOVE "Springfield" TO ws-city
           MOVE "IL" TO ws-state
           DISPLAY addr-screen
           STOP RUN.

