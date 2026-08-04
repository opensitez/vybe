*> vybe-test: cobol/move_semantics_extended/move_spaces_to_group_item_is_accepted
*> origin: languages/cobol/tests/cobol/test_move_semantics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRP.
   05 WS-A PIC X(2) VALUE "A".
   05 WS-B PIC X(2) VALUE "B".
PROCEDURE DIVISION.

    MOVE SPACES TO WS-GRP.
    STOP RUN.

