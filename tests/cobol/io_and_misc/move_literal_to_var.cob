*> vybe-test: cobol/io_and_misc/move_literal_to_var
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10).
PROCEDURE DIVISION.
    MOVE "Hello" TO X.
    STOP RUN.

