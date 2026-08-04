*> vybe-test: cobol/io_and_misc/display_multi
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(5) VALUE "Bob".
01 A PIC 9(3) VALUE 30.
PROCEDURE DIVISION.
    DISPLAY "Name: " N " Age: " A.
    STOP RUN.

