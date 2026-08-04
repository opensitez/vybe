*> vybe-test: cobol/cobol2023_nested/copy_replacing
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING ==OLD-NAME== BY ==CUST-NAME==.
    DISPLAY "After copy replacing".
    STOP RUN.

