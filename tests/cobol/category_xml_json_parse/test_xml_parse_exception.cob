*> vybe-test: cobol/category_xml_json_parse/test_xml_parse_exception
*> origin: languages/cobol/tests/cobol/test_category_xml_json_parse.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 DOC PIC X(20) VALUE '<A>1</A>'. PROCEDURE DIVISION. XML PARSE DOC PROCESSING PROCEDURE P1 ON EXCEPTION DISPLAY 'E' NOT ON EXCEPTION DISPLAY 'N'. STOP RUN. P1 SECTION. EXIT.

