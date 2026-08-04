*> vybe-test: cobol/enum_like_states/enum_state_transition_to_false_surface
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC 9 VALUE 2.
   88 STATE-NEW VALUE 1.
   88 STATE-BUSY VALUE 2.
   88 STATE-DONE VALUE 3.
PROCEDURE DIVISION.
    SET STATE-BUSY TO TRUE
    DISPLAY WS-STATE
    SET STATE-BUSY TO FALSE
    IF STATE-BUSY
        DISPLAY "BUSY"
    ELSE
        DISPLAY "NOT-BUSY"
    END-IF
    STOP RUN.

