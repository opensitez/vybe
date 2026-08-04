*> vybe-test: cobol/io_and_misc/move_num_to_var
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5).
PROCEDURE DIVISION.
    MOVE 42 TO X.
    STOP RUN.

