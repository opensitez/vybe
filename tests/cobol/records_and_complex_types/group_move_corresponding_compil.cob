*> vybe-test: cobol/records_and_complex_types/group_move_corresponding_compiles
*> origin: languages/cobol/tests/cobol/test_records_and_complex_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC.
   05 WS-NAME PIC X(10) VALUE "BOB".
   05 WS-AGE PIC 9(2) VALUE 41.
01 WS-DST.
   05 WS-NAME PIC X(10).
   05 WS-AGE PIC 9(2).
PROCEDURE DIVISION.
    MOVE CORRESPONDING WS-SRC TO WS-DST.
    STOP RUN.

