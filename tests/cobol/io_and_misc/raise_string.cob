*> vybe-test: cobol/io_and_misc/raise_string
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    RAISE EXCEPTION "Error occurred".
    STOP RUN.

