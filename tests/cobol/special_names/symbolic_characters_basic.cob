*> vybe-test: cobol/special_names/symbolic_characters_basic
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           SYMBOLIC CHARACTERS TAB IS 9
                               ESC IS 27.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-tab PIC X VALUE TAB.
       PROCEDURE DIVISION.
           DISPLAY "tab char defined"
           STOP RUN.

