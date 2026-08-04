*> vybe-test: cobol/inspect_converting/inspect_tallying_and_replacing_combined
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR ALL "L" REPLACING ALL "L" BY "R".
    STOP RUN.

