*> vybe-test: cobol/figurative_constants/spaces_to_alpha
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-name PIC X(20) VALUE "John Doe".
       PROCEDURE DIVISION.
           MOVE SPACES TO ws-name
           IF ws-name = SPACES
               DISPLAY "cleared"
           END-IF
           STOP RUN.

