*> vybe-test: cobol/level_66_78/renames_with_redefines_sibling
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 packed-data.
           05 pd-type    PIC X.
           05 pd-code    PIC 9(4).
           05 pd-amount  PIC 9(7)V99.
       66 pd-code-amount RENAMES pd-code THRU pd-amount.
       PROCEDURE DIVISION.
           MOVE "A"  TO pd-type
           MOVE 1234 TO pd-code
           MOVE 9999.99 TO pd-amount
           DISPLAY pd-code-amount
           STOP RUN.

