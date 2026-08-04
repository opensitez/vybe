*> vybe-test: cobol/special_names/class_special_chars
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS HEX-CHARS IS "0" THRU "9" "A" THRU "F" "a" THRU "f".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-hex PIC X VALUE "F".
       PROCEDURE DIVISION.
           IF ws-hex IS HEX-CHARS
               DISPLAY "valid hex"
           END-IF
           STOP RUN.

