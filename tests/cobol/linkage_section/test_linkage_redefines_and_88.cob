*> vybe-test: cobol/linkage_section/test_linkage_redefines_and_88
*> origin: languages/cobol/tests/cobol/test_linkage_section.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SUBPROG.
DATA DIVISION.
LINKAGE SECTION.
01 LNK-STATUS PIC X.
   88 IS-OK VALUE "Y".
01 LNK-STATUS-RED REDEFINES LNK-STATUS PIC 9.
PROCEDURE DIVISION USING LNK-STATUS.
    IF IS-OK
        DISPLAY "OK"
    END-IF.
    GOBACK.

