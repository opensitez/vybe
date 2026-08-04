*> vybe-test: cobol/level88_transition/level88_set_false_by_first_false_value
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SWITCH PIC X VALUE "Y".
    88 SW-ON VALUE "Y".
    88 SW-OFF VALUE "N".
PROCEDURE DIVISION.
    SET SW-ON TO FALSE.
    STOP RUN.

