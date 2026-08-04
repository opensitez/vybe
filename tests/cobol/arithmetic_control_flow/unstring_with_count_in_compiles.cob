*> vybe-test: cobol/arithmetic_control_flow/unstring_with_count_in_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(10) VALUE "AA,BBB".
01 F1 PIC X(5).
01 F2 PIC X(5).
01 C1 PIC 9(2).
01 C2 PIC 9(2).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY ","
        INTO F1 COUNT IN C1
             F2 COUNT IN C2.
    STOP RUN.

