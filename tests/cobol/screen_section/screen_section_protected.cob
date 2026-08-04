*> vybe-test: cobol/screen_section/screen_section_protected
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-readonly PIC X(20) VALUE "READ ONLY".
       SCREEN SECTION.
       01 view-screen.
           05 LINE 1 COLUMN 1 PIC X(20) FROM ws-readonly PROTECTED.
       PROCEDURE DIVISION.
           DISPLAY view-screen
           STOP RUN.

