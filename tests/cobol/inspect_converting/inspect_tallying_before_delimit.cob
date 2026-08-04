*> vybe-test: cobol/inspect_converting/inspect_tallying_before_delimiter
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(20) VALUE "HELLO WORLD".
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR ALL "L" BEFORE " ".
    STOP RUN.

