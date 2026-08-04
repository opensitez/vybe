*> vybe-test: cobol/cobol2023_json_xml/json_generate_name_override
*> origin: languages/cobol/tests/cobol/test_cobol2023_json_xml.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA.
   05 WS-FIRST-NAME PIC X(20) VALUE "Jane".
   05 WS-LAST-NAME PIC X(20) VALUE "Doe".
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-DATA
        NAME OF WS-FIRST-NAME IS "firstName"
        NAME OF WS-LAST-NAME IS "lastName".
    DISPLAY WS-JSON.
    STOP RUN.

