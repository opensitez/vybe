*> vybe-test: cobol/literals_strings_interpolation/inspect_replacing_literal_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "ABA".
PROCEDURE DIVISION.
    INSPECT S REPLACING ALL "A" BY "Z".
    STOP RUN.

