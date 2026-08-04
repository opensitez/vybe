*> vybe-test: cobol/structured_control_flow_expanded/nested_if_evaluate_mix_compiles
*> origin: languages/cobol/tests/cobol/test_structured_control_flow_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF WS-X > 0
        EVALUATE WS-X
            WHEN 1 DISPLAY "ONE"
            WHEN 2 DISPLAY "TWO"
            WHEN OTHER DISPLAY "OTHER"
        END-EVALUATE
    END-IF.
    STOP RUN.

