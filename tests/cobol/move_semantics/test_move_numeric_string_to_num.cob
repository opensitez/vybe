*> vybe-test: cobol/move_semantics/test_move_numeric_string_to_numeric_field
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(3) VALUE "123".
01 WS-DST PIC 9(3) VALUE 0.
PROCEDURE DIVISION.

    MOVE WS-SRC TO WS-DST.
    STOP RUN.

