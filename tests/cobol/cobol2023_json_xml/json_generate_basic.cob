*> vybe-test: cobol/cobol2023_json_xml/json_generate_basic
*> origin: languages/cobol/tests/cobol/test_cobol2023_json_xml.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PERSON.
   05 WS-NAME PIC X(20) VALUE "John".
   05 WS-AGE PIC 9(3) VALUE 30.
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-PERSON.
    DISPLAY WS-JSON.
    STOP RUN.

