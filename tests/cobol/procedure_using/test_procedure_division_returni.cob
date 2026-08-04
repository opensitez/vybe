*> vybe-test: cobol/procedure_using/test_procedure_division_returning
*> origin: languages/cobol/tests/cobol/test_procedure_using.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-RET PIC 9(3).
PROCEDURE DIVISION RETURNING LNK-RET.
    MOVE 100 TO LNK-RET.
    GOBACK.

