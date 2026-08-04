*> vybe-test: cobol/reference_modification/test_refmod_omitted_length
*> origin: languages/cobol/tests/cobol/test_reference_modification.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(5) VALUE "ABCDE".
PROCEDURE DIVISION.

    DISPLAY WS-TXT(3:).
    STOP RUN.

