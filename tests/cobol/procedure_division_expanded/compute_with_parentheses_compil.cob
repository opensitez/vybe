*> vybe-test: cobol/procedure_division_expanded/compute_with_parentheses_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 2.
01 WS-B PIC 9(3) VALUE 3.
01 WS-C PIC 9(3).
PROCEDURE DIVISION.
    COMPUTE WS-C = (WS-A + WS-B) * 2.
    STOP RUN.

