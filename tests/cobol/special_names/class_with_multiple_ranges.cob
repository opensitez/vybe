*> vybe-test: cobol/special_names/class_with_multiple_ranges
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CLASS ALNUM IS "0" THRU "9" "A" THRU "Z" "a" THRU "z".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-test PIC X VALUE "X".
       PROCEDURE DIVISION.
           IF ws-test IS ALNUM
               DISPLAY "alphanumeric"
           END-IF
           STOP RUN.

