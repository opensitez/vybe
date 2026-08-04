*> vybe-test: cobol/figurative_constants/space_singular_compare
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ch PIC X VALUE SPACE.
       PROCEDURE DIVISION.
           IF ws-ch = SPACE
               DISPLAY "blank"
           END-IF
           STOP RUN.

