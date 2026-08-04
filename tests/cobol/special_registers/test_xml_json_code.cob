*> vybe-test: cobol/special_registers/test_xml_json_code
*> origin: languages/cobol/tests/cobol/test_special_registers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    DISPLAY XML-CODE.
    DISPLAY JSON-CODE.
    STOP RUN.

