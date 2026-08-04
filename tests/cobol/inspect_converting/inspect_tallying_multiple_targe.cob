*> vybe-test: cobol/inspect_converting/inspect_tallying_multiple_targets
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(20) VALUE "HELLO WORLD".
01 C1 PIC 9(3) VALUE 0.
01 C2 PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C1 FOR ALL "L" C2 FOR ALL "O".
    STOP RUN.

