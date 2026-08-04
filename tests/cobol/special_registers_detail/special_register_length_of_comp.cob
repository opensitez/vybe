*> vybe-test: cobol/special_registers_detail/special_register_length_of_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
01 L PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE L = FUNCTION LENGTH(S).
    STOP RUN.

