*> vybe-test: cobol/io_and_misc/initialize_group
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 A PIC X(10) VALUE "Old".
   05 B PIC 9(5) VALUE 99.
PROCEDURE DIVISION.
    INITIALIZE REC.
    STOP RUN.

