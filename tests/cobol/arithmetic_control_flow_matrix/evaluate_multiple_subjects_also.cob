*> vybe-test: cobol/arithmetic_control_flow_matrix/evaluate_multiple_subjects_also_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE A ALSO B
        WHEN 1 ALSO 2 DISPLAY "M"
        WHEN OTHER DISPLAY "N"
    END-EVALUATE.
    STOP RUN.

