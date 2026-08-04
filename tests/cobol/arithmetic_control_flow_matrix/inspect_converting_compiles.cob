*> vybe-test: cobol/arithmetic_control_flow_matrix/inspect_converting_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(6) VALUE "abc123".
PROCEDURE DIVISION.
    INSPECT TXT CONVERTING "abc" TO "ABC".
    STOP RUN.

