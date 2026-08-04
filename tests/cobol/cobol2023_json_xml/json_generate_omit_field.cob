*> vybe-test: cobol/cobol2023_json_xml/json_generate_omit_field
*> origin: languages/cobol/tests/cobol/test_cobol2023_json_xml.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA.
   05 WS-NAME PIC X(20) VALUE "John".
   05 WS-INTERNAL PIC X(10) VALUE "secret".
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-DATA
        NAME OF WS-INTERNAL IS OMITTED.
    DISPLAY WS-JSON.
    STOP RUN.

