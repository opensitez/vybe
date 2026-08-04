*> vybe-test: cobol/literals_strings_interpolation/inspect_tallying_literal_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "ABA".
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR ALL "A".
    STOP RUN.

