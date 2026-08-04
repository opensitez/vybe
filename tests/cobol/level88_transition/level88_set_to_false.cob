*> vybe-test: cobol/level88_transition/level88_set_to_false
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "Y".
    88 FLAG-ON VALUE "Y".
    88 FLAG-OFF VALUE "N".
PROCEDURE DIVISION.
    SET FLAG-ON TO FALSE.
    STOP RUN.

