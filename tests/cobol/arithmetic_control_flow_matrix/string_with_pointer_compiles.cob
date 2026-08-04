*> vybe-test: cobol/arithmetic_control_flow_matrix/string_with_pointer_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(2) VALUE "AB".
01 B PIC X(2) VALUE "CD".
01 R PIC X(10).
01 P PIC 9(2) VALUE 1.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R WITH POINTER P.
    STOP RUN.

