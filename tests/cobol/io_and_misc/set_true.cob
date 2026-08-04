*> vybe-test: cobol/io_and_misc/set_true
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1).
   88 IS-ON VALUE 1.
PROCEDURE DIVISION.
    SET IS-ON TO TRUE.
    STOP RUN.

