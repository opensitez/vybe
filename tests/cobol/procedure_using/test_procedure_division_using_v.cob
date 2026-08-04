*> vybe-test: cobol/procedure_using/test_procedure_division_using_value
*> origin: languages/cobol/tests/cobol/test_procedure_using.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-VAL PIC 9(3).
PROCEDURE DIVISION USING BY VALUE LNK-VAL.
    DISPLAY LNK-VAL.
    GOBACK.

