*> vybe-test: cobol/local_storage/test_local_storage_value_clause
*> origin: languages/cobol/tests/cobol/test_local_storage.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. LS-PROG.
DATA DIVISION.
LOCAL-STORAGE SECTION.
01 LS-X PIC 9(3) VALUE ZERO.
01 LS-Y PIC X(5) VALUE SPACES.
PROCEDURE DIVISION.
    DISPLAY LS-X.
    DISPLAY LS-Y.
    GOBACK.

