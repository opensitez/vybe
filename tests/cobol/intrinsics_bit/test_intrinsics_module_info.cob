*> vybe-test: cobol/intrinsics_bit/test_intrinsics_module_info
*> origin: languages/cobol/tests/cobol/test_intrinsics_bit.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(30).
PROCEDURE DIVISION.

    MOVE FUNCTION MODULE-NAME TO WS-NAME.
    STOP RUN.

