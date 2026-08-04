*> vybe-test: cobol/io_and_misc/move_corresponding
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 WS-NAME PIC X(10) VALUE "Alice".
   05 WS-AGE PIC 9(3) VALUE 30.
01 DST.
   05 WS-NAME PIC X(10).
   05 WS-AGE PIC 9(3).
PROCEDURE DIVISION.
    MOVE CORRESPONDING SRC TO DST.
    STOP RUN.

