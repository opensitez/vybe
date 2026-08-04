*> vybe-test: cobol/enum_like_states/enum_state_transition_runtime
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
    SET STATE-NEW TO TRUE
    IF STATE-NEW
        DISPLAY "NEW"
    END-IF
    SET STATE-RUN TO TRUE
    IF STATE-RUN
        DISPLAY "RUN"
    END-IF
    STOP RUN.

