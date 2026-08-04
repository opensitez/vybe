*> vybe-test: cobol/arithmetic_control_flow_matrix/perform_with_test_after_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM WITH TEST AFTER UNTIL I > 2
        ADD 1 TO I
    END-PERFORM.
    STOP RUN.

