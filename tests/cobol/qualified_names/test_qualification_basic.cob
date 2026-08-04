*> vybe-test: cobol/qualified_names/test_qualification_basic
*> origin: languages/cobol/tests/cobol/test_qualified_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GROUP-A.
   05 WS-NAME PIC X(5) VALUE "ALICE".
01 WS-GROUP-B.
   05 WS-NAME PIC X(5) VALUE "BOB  ".
PROCEDURE DIVISION.

    DISPLAY WS-NAME IN WS-GROUP-A.
    DISPLAY WS-NAME OF WS-GROUP-B.
    STOP RUN.

