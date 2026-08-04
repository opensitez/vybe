*> vybe-test: cobol/move_semantics/test_move_high_values_to_alpha
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DST PIC X(5).
PROCEDURE DIVISION.

    MOVE HIGH-VALUES TO WS-DST.
    STOP RUN.

