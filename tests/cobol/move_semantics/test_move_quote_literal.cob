*> vybe-test: cobol/move_semantics/test_move_quote_literal
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DST PIC X(1).
PROCEDURE DIVISION.

    MOVE QUOTE TO WS-DST.
    STOP RUN.

