*> vybe-test: cobol/arithmetic_control_flow_matrix/unstring_with_pointer_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(12) VALUE "AA,BBB,CC".
01 F1 PIC X(5).
01 F2 PIC X(5).
01 P PIC 9(2) VALUE 1.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2 WITH POINTER P.
    STOP RUN.

