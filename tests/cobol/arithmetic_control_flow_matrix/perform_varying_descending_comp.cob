*> vybe-test: cobol/arithmetic_control_flow_matrix/perform_varying_descending_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 5 BY -1 UNTIL I < 1
        DISPLAY I
    END-PERFORM.
    STOP RUN.

