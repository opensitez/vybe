*> vybe-test: cobol/io_and_misc/call_multi_args
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10).
01 B PIC 9(5).
PROCEDURE DIVISION.
    CALL "PROCESS" USING A B.
    STOP RUN.

