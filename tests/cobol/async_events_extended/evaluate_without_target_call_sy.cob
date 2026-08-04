*> vybe-test: cobol/async_events_extended/evaluate_without_target_call_syntax_still_compiles
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-YN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1
            CALL "SUB" ON EXCEPTION DISPLAY "ONE" NOT ON EXCEPTION DISPLAY "ONE-OK" END-CALL
        WHEN 2
            CALL "SUB" ON EXCEPTION DISPLAY "TWO" NOT ON EXCEPTION DISPLAY "TWO-OK" END-CALL
        WHEN OTHER
            CALL "SUB" ON EXCEPTION DISPLAY "OTHER" NOT ON EXCEPTION DISPLAY "OTHER-OK" END-CALL
    END-EVALUATE
    STOP RUN.

