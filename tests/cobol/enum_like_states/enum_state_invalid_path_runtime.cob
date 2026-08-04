*> vybe-test: cobol/enum_like_states/enum_state_invalid_path_runtime
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC 9 VALUE 9.
   88 STATE-NEW VALUE 1.
   88 STATE-DONE VALUE 2.
PROCEDURE DIVISION.
    IF NOT STATE-NEW
       AND NOT STATE-DONE
        DISPLAY "UNKNOWN"
    END-IF
    STOP RUN.

