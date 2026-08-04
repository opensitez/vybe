*> vybe-test: cobol/arithmetic_control_flow_matrix/call_nested_in_if_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF F = 1
        CALL "OKM"
    ELSE
        CALL "NOM"
    END-IF.
    STOP RUN.

