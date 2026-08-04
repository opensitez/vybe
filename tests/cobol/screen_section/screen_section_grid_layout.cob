*> vybe-test: cobol/screen_section/screen_section_grid_layout
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-vals.
           05 ws-val-1 PIC 999 VALUE 100.
           05 ws-val-2 PIC 999 VALUE 200.
           05 ws-val-3 PIC 999 VALUE 300.
       SCREEN SECTION.
       01 grid-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1  VALUE "Col1".
           05 LINE 1 COLUMN 10 VALUE "Col2".
           05 LINE 1 COLUMN 20 VALUE "Col3".
           05 LINE 2 COLUMN 1  PIC 999 FROM ws-val-1.
           05 LINE 2 COLUMN 10 PIC 999 FROM ws-val-2.
           05 LINE 2 COLUMN 20 PIC 999 FROM ws-val-3.
       PROCEDURE DIVISION.
           DISPLAY grid-screen
           STOP RUN.

