*> vybe-test: cobol/figurative_constants/all_in_string
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pad PIC X(5).
       01 ws-result PIC X(15).
       PROCEDURE DIVISION.
           MOVE ALL "." TO ws-pad
           STRING "Hi" DELIMITED SIZE
                  ws-pad DELIMITED SIZE
                  INTO ws-result
           DISPLAY ws-result
           STOP RUN.

