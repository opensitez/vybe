*> vybe-test: cobol/arithmetic_control_flow/inspect_replacing_first_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(6) VALUE "AAAAAA".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING FIRST "A" BY "B".
    STOP RUN.

