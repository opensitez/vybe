*> vybe-test: cobol/io_and_misc/display_var
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10) VALUE "Test".
PROCEDURE DIVISION.
    DISPLAY X.
    STOP RUN.

