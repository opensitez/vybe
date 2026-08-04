*> vybe-test: cobol/local_storage/test_local_storage_basics
*> origin: languages/cobol/tests/cobol/test_local_storage.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. LS-PROG.
DATA DIVISION.
LOCAL-STORAGE SECTION.
01 LS-NUM PIC 9(3) VALUE 100.
01 LS-STR PIC X(5) VALUE "HELLO".
01 LS-TABLE.
   05 LS-ITEM OCCURS 3 TIMES PIC 9(3).
PROCEDURE DIVISION.
    ADD 1 TO LS-NUM.
    DISPLAY LS-NUM.
    DISPLAY LS-STR.
    GOBACK.

