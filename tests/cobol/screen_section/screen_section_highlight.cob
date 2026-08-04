*> vybe-test: cobol/screen_section/screen_section_highlight
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC X(10).
       SCREEN SECTION.
       01 hi-screen.
           05 LINE 1 COLUMN 1 VALUE "Field: " HIGHLIGHT.
           05 LINE 1 COLUMN 8 PIC X(10) USING ws-val HIGHLIGHT.
       PROCEDURE DIVISION.
           MOVE "test" TO ws-val
           DISPLAY hi-screen
           STOP RUN.

