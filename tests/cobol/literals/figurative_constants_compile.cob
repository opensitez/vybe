*> vybe-test: cobol/literals/figurative_constants_compile
*> origin: languages/cobol/tests/cobol/test_literals.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FILL PIC X(5) VALUE SPACES.
01 WS-ZERO PIC 9(5) VALUE ZERO.
PROCEDURE DIVISION.
    MOVE ZEROS TO WS-ZERO.
    MOVE SPACES TO WS-FILL.
    STOP RUN.

