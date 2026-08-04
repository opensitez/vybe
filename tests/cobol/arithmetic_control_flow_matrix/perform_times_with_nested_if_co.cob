*> vybe-test: cobol/arithmetic_control_flow_matrix/perform_times_with_nested_if_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM 3 TIMES
        ADD 1 TO I
        IF I = 2 DISPLAY "M" END-IF
    END-PERFORM.
    STOP RUN.

