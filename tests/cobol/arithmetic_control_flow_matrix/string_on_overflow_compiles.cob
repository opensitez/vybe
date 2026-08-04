*> vybe-test: cobol/arithmetic_control_flow_matrix/string_on_overflow_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "ABCDE".
01 B PIC X(5) VALUE "FGHIJ".
01 R PIC X(3).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R ON OVERFLOW DISPLAY "OV" END-STRING.
    STOP RUN.

