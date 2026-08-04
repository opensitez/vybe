*> vybe-test: cobol/qualified_names/test_qualification_statement_usage
*> origin: languages/cobol/tests/cobol/test_qualified_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GROUP-A.
   05 WS-VAL PIC 9(3) VALUE 10.
01 WS-GROUP-B.
   05 WS-VAL PIC 9(3) VALUE 20.
PROCEDURE DIVISION.

    ADD WS-VAL IN WS-GROUP-A TO WS-VAL IN WS-GROUP-B.
    STOP RUN.

