*> vybe-test: cobol/io_and_misc/move_spaces
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10).
PROCEDURE DIVISION.
    MOVE SPACES TO X.
    STOP RUN.

