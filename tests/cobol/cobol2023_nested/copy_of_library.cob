*> vybe-test: cobol/cobol2023_nested/copy_of_library
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    COPY DATE-UTILS OF COMMON-LIB.
    DISPLAY "Copy from library".
    STOP RUN.

