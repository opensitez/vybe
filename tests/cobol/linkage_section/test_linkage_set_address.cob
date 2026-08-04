*> vybe-test: cobol/linkage_section/test_linkage_set_address
*> origin: languages/cobol/tests/cobol/test_linkage_section.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR POINTER.
LINKAGE SECTION.
01 LNK-ITEM PIC X(10).
PROCEDURE DIVISION.
    SET ADDRESS OF LNK-ITEM TO WS-PTR.
    GOBACK.

