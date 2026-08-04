*> vybe-test: cobol/new_features/invoke_basic
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 OBJ PIC X(10).
01 RES PIC X(10).
PROCEDURE DIVISION.
    INVOKE OBJ GET-NAME RETURNING RES.
    STOP RUN.

