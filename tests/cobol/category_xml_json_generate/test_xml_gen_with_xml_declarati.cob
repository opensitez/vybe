*> vybe-test: cobol/category_xml_json_generate/test_xml_gen_with_xml_declaration
*> origin: languages/cobol/tests/cobol/test_category_xml_json_generate.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC 9 VALUE 1. 01 DOC PIC X(20). PROCEDURE DIVISION. XML GENERATE DOC FROM R WITH XML-DECLARATION. DISPLAY 'OK'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

