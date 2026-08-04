*> vybe-test: cobol/special_names/class_digit_chars
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS DIGIT-CHARS IS "0" THRU "9".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-char PIC X VALUE "5".
       PROCEDURE DIVISION.
           IF ws-char IS DIGIT-CHARS
               DISPLAY "is digit"
           ELSE
               DISPLAY "not digit"
           END-IF
           STOP RUN.

