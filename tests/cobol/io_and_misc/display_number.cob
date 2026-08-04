*> vybe-test: cobol/io_and_misc/display_number
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5) VALUE 42.
PROCEDURE DIVISION.
    DISPLAY X.
    STOP RUN.

