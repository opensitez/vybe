*> vybe-test: cobol/arithmetic_control_flow_matrix/inspect_replacing_all_letters_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(6) VALUE "ABCABC".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING ALL "A" BY "Z".
    STOP RUN.

