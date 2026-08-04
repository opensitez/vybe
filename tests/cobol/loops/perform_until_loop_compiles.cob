*> vybe-test: cobol/loops/perform_until_loop_compiles
*> origin: languages/cobol/tests/cobol/test_loops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-COUNT >= 3
        ADD 1 TO WS-COUNT
    END-PERFORM.
    STOP RUN.

