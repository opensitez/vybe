*> vybe-test: cobol/special_names/symbolic_characters_multiple
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           SYMBOLIC CHARACTERS NULL-CHAR IS 1
                               BELL      IS 7
                               CR-CHAR   IS 13
                               LF-CHAR   IS 10.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-cr PIC X VALUE CR-CHAR.
       PROCEDURE DIVISION.
           DISPLAY "control chars defined"
           STOP RUN.

