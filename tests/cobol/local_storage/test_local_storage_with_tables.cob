*> vybe-test: cobol/local_storage/test_local_storage_with_tables
*> origin: languages/cobol/tests/cobol/test_local_storage.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. LS-TABLE.
DATA DIVISION.
LOCAL-STORAGE SECTION.
01 LS-LIMIT PIC 9(2) VALUE 2.
01 LS-ENTRIES.
   05 LS-ENTRY OCCURS 2 TIMES PIC X(3) VALUE "A".
PROCEDURE DIVISION.
    DISPLAY LS-LIMIT.
    DISPLAY LS-ENTRY(1).
    GOBACK.

