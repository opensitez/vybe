*> vybe-test: cobol/cobol/json_parse
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. JSONPAR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-JSON PIC X(100) VALUE '{"name":"Alice","age":30}'.
01 WS-PERSON.
   05 WS-NAME PIC X(10).
   05 WS-AGE  PIC 9(3).
PROCEDURE DIVISION.
    JSON PARSE WS-JSON INTO WS-PERSON.
    DISPLAY WS-NAME.
    STOP RUN.

