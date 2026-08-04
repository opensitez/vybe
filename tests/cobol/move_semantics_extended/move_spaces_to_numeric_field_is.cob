*> vybe-test: cobol/move_semantics_extended/move_spaces_to_numeric_field_is_accepted
*> origin: languages/cobol/tests/cobol/test_move_semantics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(3) VALUE 123.
PROCEDURE DIVISION.

    MOVE SPACES TO WS-NUM.
    STOP RUN.

