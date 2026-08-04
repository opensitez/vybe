*> vybe-test: cobol/new_features/invoke_using
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 OBJ PIC X(10).
01 ARG PIC X(10) VALUE "Test".
PROCEDURE DIVISION.
    INVOKE OBJ PROCESS USING ARG.
    STOP RUN.

