*> vybe-test: cobol/rounded_modes/rounded_mode_nearest_toward_zero_pos
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE NEAREST-TOWARD-ZERO = 2.5
           DISPLAY ws-result
           STOP RUN.

