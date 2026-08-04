*> vybe-test: cobol/screen_section/screen_section_reverse_video
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-title PIC X(20) VALUE "STATUS: OK".
       SCREEN SECTION.
       01 status-bar.
           05 LINE 24 COLUMN 1 PIC X(20) FROM ws-title REVERSE-VIDEO.
       PROCEDURE DIVISION.
           DISPLAY status-bar
           STOP RUN.

