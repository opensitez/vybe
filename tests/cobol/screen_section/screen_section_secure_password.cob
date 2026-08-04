*> vybe-test: cobol/screen_section/screen_section_secure_password
*> origin: languages/cobol/tests/cobol/test_screen_section.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pass PIC X(16).
       SCREEN SECTION.
       01 login-screen.
           05 BLANK SCREEN.
           05 LINE 5 COLUMN 20 VALUE "Password: ".
           05 LINE 5 COLUMN 30 PIC X(16) USING ws-pass SECURE.
       PROCEDURE DIVISION.
           DISPLAY login-screen
           STOP RUN.

