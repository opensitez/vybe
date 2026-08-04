*> vybe-test: cobol/special_names/console_is_crt
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CONSOLE IS CRT.
       PROCEDURE DIVISION.
           DISPLAY "console output"
           STOP RUN.

