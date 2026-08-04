*> vybe-test: cobol/enum_like_states/evaluate_on_enum_like_state_compiles
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATE PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE WS-STATE
        WHEN 1 DISPLAY "NEW"
        WHEN 2 DISPLAY "RUN"
        WHEN 3 DISPLAY "DONE"
        WHEN OTHER DISPLAY "UNK"
    END-EVALUATE.
    STOP RUN.

