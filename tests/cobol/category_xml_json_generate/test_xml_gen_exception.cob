*> vybe-test: cobol/category_xml_json_generate/test_xml_gen_exception
*> origin: languages/cobol/tests/cobol/test_category_xml_json_generate.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R. 05 A PIC 9 VALUE 1. 01 DOC PIC X(2). PROCEDURE DIVISION. XML GENERATE DOC FROM R ON EXCEPTION DISPLAY 'E' NOT ON EXCEPTION DISPLAY 'N'. STOP RUN.

