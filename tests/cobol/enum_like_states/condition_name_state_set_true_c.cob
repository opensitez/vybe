*> vybe-test: cobol/enum_like_states/condition_name_state_set_true_compiles
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC 9 VALUE 0.
   88 STATE-NEW VALUE 1.
   88 STATE-DONE VALUE 2.
PROCEDURE DIVISION.
    SET STATE-NEW TO TRUE.
    IF STATE-NEW
        DISPLAY "NEW"
    END-IF.
    STOP RUN.

