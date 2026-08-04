*> vybe-test: cobol/category_xml_json_parse/test_json_parse_exception
*> origin: languages/cobol/tests/cobol/test_category_xml_json_parse.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 DOC PIC X(20) VALUE '{"A":1}'. 01 R. 05 A PIC 9. PROCEDURE DIVISION. JSON PARSE DOC INTO R ON EXCEPTION DISPLAY 'E' NOT ON EXCEPTION DISPLAY 'N'. STOP RUN.

