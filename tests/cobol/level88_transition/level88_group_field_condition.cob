*> vybe-test: cobol/level88_transition/level88_group_field_condition
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RECORD-STATUS PIC X VALUE "A".
    88 RECORD-ACTIVE VALUE "A".
    88 RECORD-DELETED VALUE "D".
PROCEDURE DIVISION.
    IF RECORD-ACTIVE
        DISPLAY "ACTIVE"
    END-IF.
    STOP RUN.

