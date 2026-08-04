*> vybe-test: cobol/procedure_using/test_procedure_division_with_returning_and_using
*> origin: languages/cobol/tests/cobol/test_procedure_using.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-IN PIC 9(3) VALUE 0.
01 LNK-OUT PIC 9(3).
PROCEDURE DIVISION USING LNK-IN RETURNING LNK-OUT.
    COMPUTE LNK-OUT = LNK-IN + 1.
    GOBACK.

