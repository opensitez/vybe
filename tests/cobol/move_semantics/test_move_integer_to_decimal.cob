*> vybe-test: cobol/move_semantics/test_move_integer_to_decimal
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC 9(3) VALUE 123.
01 WS-DST PIC 9(3)V99 VALUE 0.0.
PROCEDURE DIVISION.

    MOVE WS-SRC TO WS-DST.
    STOP RUN.

