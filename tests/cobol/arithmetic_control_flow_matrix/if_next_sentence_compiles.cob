*> vybe-test: cobol/arithmetic_control_flow_matrix/if_next_sentence_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF A = 1
        CONTINUE
    ELSE
        CONTINUE
    END-IF.
    STOP RUN.

