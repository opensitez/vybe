*> vybe-test: cobol/move_group_redefines/redefines_condition_name_on_redefine
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SWITCH PIC X VALUE "Y".
01 SWITCH-NUM REDEFINES SWITCH PIC 9.
    88 SWITCH-ON VALUE 1.
PROCEDURE DIVISION.
    IF SWITCH-ON
        DISPLAY "ON"
    ELSE
        DISPLAY "OFF"
    END-IF.
    STOP RUN.

