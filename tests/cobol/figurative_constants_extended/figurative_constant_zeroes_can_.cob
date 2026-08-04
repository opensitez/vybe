*> vybe-test: cobol/figurative_constants_extended/figurative_constant_zeroes_can_be_moved_into_group_item
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRP.
   05 WS-A PIC 9(3) VALUE 7.
   05 WS-B PIC 9(3) VALUE 8.
PROCEDURE DIVISION.

    MOVE ZEROS TO WS-GRP.
    STOP RUN.

