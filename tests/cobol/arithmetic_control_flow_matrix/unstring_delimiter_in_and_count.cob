*> vybe-test: cobol/arithmetic_control_flow_matrix/unstring_delimiter_in_and_count_in_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(12) VALUE "AA,BBB".
01 F1 PIC X(5).
01 D1 PIC X.
01 C1 PIC 9(2).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 DELIMITER IN D1 COUNT IN C1.
    STOP RUN.

