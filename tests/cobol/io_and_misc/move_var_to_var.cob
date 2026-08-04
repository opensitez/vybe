*> vybe-test: cobol/io_and_misc/move_var_to_var
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10) VALUE "Hi".
01 B PIC X(10).
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

