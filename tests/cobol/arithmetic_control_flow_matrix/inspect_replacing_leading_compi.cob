*> vybe-test: cobol/arithmetic_control_flow_matrix/inspect_replacing_leading_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(6) VALUE "000123".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING LEADING "0" BY " ".
    STOP RUN.

