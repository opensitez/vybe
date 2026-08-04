*> vybe-test: cobol/io_and_misc/set_false
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1).
   88 IS-OFF VALUE 0.
PROCEDURE DIVISION.
    SET IS-OFF TO FALSE.
    STOP RUN.

