*> vybe-test: cobol/pointers/pointer_variable_declares
*> origin: languages/cobol/tests/cobol/test_pointers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR USAGE POINTER.
01 WS-VAL PIC X(5) VALUE "DATA".
PROCEDURE DIVISION.
    SET WS-PTR TO NULL.
    STOP RUN.

