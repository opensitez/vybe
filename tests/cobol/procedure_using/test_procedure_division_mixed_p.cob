*> vybe-test: cobol/procedure_using/test_procedure_division_mixed_passing_modes
*> origin: languages/cobol/tests/cobol/test_procedure_using.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-VALUE PIC 9(3).
01 LNK-TEXT PIC X(5).
PROCEDURE DIVISION USING BY VALUE LNK-VALUE BY REFERENCE LNK-TEXT.
    DISPLAY LNK-VALUE.
    DISPLAY LNK-TEXT.
    GOBACK.

