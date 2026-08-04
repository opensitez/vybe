*> vybe-test: cobol/procedure_using/test_procedure_division_using
*> origin: languages/cobol/tests/cobol/test_procedure_using.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-PARAM1 PIC 9(3).
01 LNK-PARAM2 PIC X(5).
PROCEDURE DIVISION USING LNK-PARAM1 LNK-PARAM2.
    DISPLAY LNK-PARAM1.
    DISPLAY LNK-PARAM2.
    GOBACK.

