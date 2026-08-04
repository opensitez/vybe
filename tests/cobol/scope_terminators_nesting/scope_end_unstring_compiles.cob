*> vybe-test: cobol/scope_terminators_nesting/scope_end_unstring_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(10) VALUE "A,B,C".
01 F1 PIC X(5).
01 F2 PIC X(5).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2
    END-UNSTRING.
    STOP RUN.

