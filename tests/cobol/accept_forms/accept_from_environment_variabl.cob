*> vybe-test: cobol/accept_forms/accept_from_environment_variable
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ENV-VAL PIC X(80).
PROCEDURE DIVISION.
    ACCEPT ENV-VAL FROM ENVIRONMENT "PATH".
    STOP RUN.

