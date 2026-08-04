*> vybe-test: cobol/arithmetic_control_flow_matrix/evaluate_when_any_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN 5 ALSO ANY DISPLAY "HIT"
        WHEN OTHER DISPLAY "MISS"
    END-EVALUATE.
    STOP RUN.

