*> vybe-test: cobol/scope_terminators_nesting/scope_end_string_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 B PIC X(5) VALUE "WORLD".
01 R PIC X(20).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R
    END-STRING.
    STOP RUN.

