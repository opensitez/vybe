*> vybe-test: cobol/move_semantics_extended/move_zeroes_to_group_item_is_accepted
*> origin: languages/cobol/tests/cobol/test_move_semantics_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRP.
   05 WS-A PIC 9(2) VALUE 1.
   05 WS-B PIC 9(2) VALUE 2.
PROCEDURE DIVISION.

    MOVE ZEROES TO WS-GRP.
    STOP RUN.

