*> vybe-test: cobol/enum_like_states/condition_name_state_transition_compiles
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC 9 VALUE 0.
   88 STATE-NEW VALUE 1.
   88 STATE-RUN VALUE 2.
   88 STATE-DONE VALUE 3.
PROCEDURE DIVISION.
    SET STATE-NEW TO TRUE.
    SET STATE-RUN TO TRUE.
    SET STATE-DONE TO TRUE.
    IF STATE-DONE
        DISPLAY "DONE"
    END-IF.
    STOP RUN.

