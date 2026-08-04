*> vybe-test: cobol/figurative_constants_extended/figurative_constant_spaces_can_be_moved_into_group_item
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRP.
   05 WS-A PIC X(2) VALUE "AA".
   05 WS-B PIC X(2) VALUE "BB".
PROCEDURE DIVISION.

    MOVE SPACES TO WS-GRP.
    STOP RUN.

