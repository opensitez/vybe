*> vybe-test: cobol/arithmetic_control_flow/set_index_up_and_down_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TAB PIC 9 OCCURS 5 TIMES INDEXED BY I.
PROCEDURE DIVISION.
    SET I TO 1.
    SET I UP BY 2.
    SET I DOWN BY 1.
    STOP RUN.

