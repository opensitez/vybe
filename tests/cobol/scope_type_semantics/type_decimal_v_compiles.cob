*> vybe-test: cobol/scope_type_semantics/type_decimal_v_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V99.
PROCEDURE DIVISION.
    MOVE 123.45 TO A.
    STOP RUN.

