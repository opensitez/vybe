*> vybe-test: cobol/cobol2023_json_xml/json_parse_basic
*> origin: languages/cobol/tests/cobol/test_cobol2023_json_xml.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-JSON PIC X(100) VALUE '{"name":"Alice","age":25}'.
01 WS-DATA.
   05 WS-NAME PIC X(20).
   05 WS-AGE PIC 9(3).
PROCEDURE DIVISION.
    JSON PARSE WS-JSON INTO WS-DATA.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.

