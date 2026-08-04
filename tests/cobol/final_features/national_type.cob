*> vybe-test: cobol/final_features/national_type
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. UTF8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(30) NATIONAL VALUE "Unicode Test".
PROCEDURE DIVISION.
    DISPLAY WS-NAME.
    STOP RUN.

