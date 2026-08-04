*> vybe-test: cobol/cobol2023_string_ops/unstring_multi_delimiter
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA PIC X(30) VALUE "John,Doe;30".
01 WS-FIRST PIC X(10).
01 WS-LAST PIC X(10).
01 WS-AGE PIC X(5).
PROCEDURE DIVISION.
    UNSTRING WS-DATA
        DELIMITED BY "," OR ";"
        INTO WS-FIRST WS-LAST WS-AGE.
    DISPLAY WS-FIRST.
    DISPLAY WS-LAST.
    DISPLAY WS-AGE.
    STOP RUN.

