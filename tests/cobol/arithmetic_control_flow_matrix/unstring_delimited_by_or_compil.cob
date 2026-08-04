*> vybe-test: cobol/arithmetic_control_flow_matrix/unstring_delimited_by_or_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(12) VALUE "A,B;C".
01 F1 PIC X(3).
01 F2 PIC X(3).
01 F3 PIC X(3).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," OR ";" INTO F1 F2 F3.
    STOP RUN.

