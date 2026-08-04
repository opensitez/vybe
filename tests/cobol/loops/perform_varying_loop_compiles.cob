*> vybe-test: cobol/loops/perform_varying_loop_compiles
*> origin: languages/cobol/tests/cobol/test_loops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-COUNT FROM 1 BY 2 UNTIL WS-COUNT > 5
        DISPLAY WS-COUNT
    END-PERFORM.
    STOP RUN.

